// ============================================================
//  Netlify Function: /.netlify/functions/bookings   (Microsoft Graph)
//
//  GET  ?room=<email>&date=YYYY-MM-DD  -> { bookings:[{start,end,name}] }
//  POST { room,name,email,date,start,end,attendees,purpose }
//         -> 200 { ok, ref }  |  409 { error:"conflict" }
//
//  Blocking strategy (race-proof without a database):
//    1. Pre-check the room's calendar for the window  -> fast 409 on obvious clash.
//    2. Create the event on the ROOM mailbox calendar.
//    3. Re-read the window. If more than one event overlaps, the racers tie-break
//       deterministically by createdDateTime (then id); every loser deletes its own
//       event and returns 409. Exactly one survives. No double-booking possible.
//    4. Send a branded confirmation from bookings@ (Mail.Send).
//
//  The room is also set to AutoAccept / AllowConflicts:$false in Exchange, so any
//  booking made the normal Outlook way is blocked too.
// ============================================================

const TENANT_ID = process.env.TENANT_ID || "0ccd0789-fd2c-48ff-9f4f-4b2bf5af070a";
const CLIENT_ID = process.env.CLIENT_ID || "71901e39-aa4d-4f67-a78f-707576c15925";
const CLIENT_SECRET = process.env.CLIENT_SECRET;           // <-- env only, never in code
// Address the room by its directory object ID (GUID), not its SMTP address:
// room mailbox UPNs can differ from the email and make /users/{smtp} throw
// ErrorInvalidUser. The GUID is unambiguous.
const ROOM_EMAIL   = process.env.ROOM_EMAIL   || "8d2de945-c4e3-411f-916f-8a4e3788db0d";
const ROOM_LABEL   = process.env.ROOM_LABEL   || "Area3_HQ-BoardMeeting@systemrapid.com";
const SENDER_EMAIL = process.env.SENDER_EMAIL || "bookings@systemrapid.com";
const ADMIN_EMAIL  = process.env.ADMIN_EMAIL  || "bookings@systemrapid.com";
const ORG_NAME     = process.env.ORG_NAME     || "FCC–Almabani Joint Venture";
const TZ_WINDOWS   = "Arab Standard Time";   // KSA, fixed UTC+03:00
const TZ_OFFSET    = "+03:00";
const GRAPH = "https://graph.microsoft.com/v1.0";

function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      "content-type": "application/json",
      "access-control-allow-origin": "*",
      "access-control-allow-methods": "GET,POST,OPTIONS",
      "access-control-allow-headers": "content-type",
    },
    body: JSON.stringify(body),
  };
}

// ---- token (client credentials) with simple in-container cache ----
let _tok = { value: null, exp: 0 };
async function getToken() {
  if (_tok.value && Date.now() < _tok.exp - 60000) return _tok.value;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body,
  });
  const data = await r.json();
  if (!r.ok) throw new Error("token: " + (data.error_description || r.status));
  _tok = { value: data.access_token, exp: Date.now() + data.expires_in * 1000 };
  return _tok.value;
}

async function graph(method, path, { body, prefer } = {}) {
  const token = await getToken();
  const headers = { authorization: `Bearer ${token}`, "content-type": "application/json" };
  if (prefer) headers["Prefer"] = prefer;
  const r = await fetch(GRAPH + path, {
    method, headers, body: body ? JSON.stringify(body) : undefined,
  });
  const text = await r.text();
  const data = text ? JSON.parse(text) : {};
  if (!r.ok) {
    const err = new Error(`graph ${method} ${path}: ${r.status}`);
    err.status = r.status; err.data = data;
    throw err;
  }
  return data;
}

const hm = (dt) => String(dt).slice(11, 16); // "....T09:00:00..." -> "09:00"

async function listWindow(room, startLocal, endLocal) {
  // events on the room calendar overlapping [startLocal, endLocal)
  const q = `?startDateTime=${encodeURIComponent(startLocal + TZ_OFFSET)}`
          + `&endDateTime=${encodeURIComponent(endLocal + TZ_OFFSET)}`
          + `&$select=id,subject,start,end,createdDateTime,isCancelled`
          + `&$orderby=start/dateTime&$top=100`;
  const data = await graph("GET", `/users/${encodeURIComponent(room)}/calendarView${q}`,
    { prefer: `outlook.timezone="${TZ_WINDOWS}"` });
  return (data.value || []).filter(e => !e.isCancelled);
}

const overlaps = (aS, aE, bS, bE) => aS < bE && aE > bS;

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return json(200, {});
  if (!CLIENT_SECRET) return json(500, { error: "server not configured (CLIENT_SECRET missing)" });

  try {
    // -------------------- availability --------------------
    if (event.httpMethod === "GET") {
      const q = event.queryStringParameters || {};
      const room = ROOM_EMAIL; // pilot: single room, always the GUID
      const date = q.date;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date || "")) return json(400, { error: "date required (YYYY-MM-DD)" });
      const evs = await listWindow(room, `${date}T00:00:00`, `${date}T23:59:59`);
      const bookings = evs.map(e => ({
        start: hm(e.start.dateTime),
        end: hm(e.end.dateTime),
        name: (e.subject || "Booked").split(" — ")[0],
      }));
      return json(200, { bookings });
    }

    // -------------------- create booking --------------------
    if (event.httpMethod === "POST") {
      let b;
      try { b = JSON.parse(event.body || "{}"); } catch { return json(400, { error: "invalid JSON" }); }
      const room = ROOM_EMAIL; // pilot: single room, always the GUID
      for (const k of ["name", "email", "date", "start", "end"]) {
        if (!b[k]) return json(400, { error: `missing field: ${k}` });
      }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(b.date) ||
          !/^\d{2}:\d{2}$/.test(b.start) || !/^\d{2}:\d{2}$/.test(b.end)) {
        return json(400, { error: "bad date or time format" });
      }
      if (!(b.start < b.end)) return json(400, { error: "end must be after start" });

      const startLocal = `${b.date}T${b.start}:00`;
      const endLocal   = `${b.date}T${b.end}:00`;
      const ref = "BK-" + Math.random().toString(36).toUpperCase().slice(2, 8);

      // 1) fast pre-check
      const before = await listWindow(room, startLocal, endLocal);
      if (before.some(e => overlaps(e.start.dateTime, e.end.dateTime, startLocal, endLocal))) {
        return json(409, { error: "conflict" });
      }

      // 2) create on the room calendar
      const created = await graph("POST", `/users/${encodeURIComponent(room)}/events`, {
        body: {
          subject: `${b.name} — ${b.purpose ? b.purpose : "Meeting"}  [${ref}]`,
          body: {
            contentType: "HTML",
            content: `Booked via the room booking portal.<br>Ref: ${ref}<br>`
              + `Booker: ${esc(b.name)} &lt;${esc(b.email)}&gt;<br>`
              + `Attendees: ${esc(b.attendees || "—")}<br>Purpose: ${esc(b.purpose || "—")}`,
          },
          start: { dateTime: startLocal, timeZone: TZ_WINDOWS },
          end:   { dateTime: endLocal,   timeZone: TZ_WINDOWS },
          location: { displayName: ROOM_LABEL },
          transactionId: ref,
        },
      });

      // 3) verify — deterministic loser-deletes
      const after = await listWindow(room, startLocal, endLocal);
      const clashing = after.filter(e =>
        overlaps(e.start.dateTime, e.end.dateTime, startLocal, endLocal));
      if (clashing.length > 1) {
        clashing.sort((x, y) => {
          const c = String(x.createdDateTime).localeCompare(String(y.createdDateTime));
          return c !== 0 ? c : String(x.id).localeCompare(String(y.id));
        });
        if (clashing[0].id !== created.id) {
          // we lost the race — remove our event, report conflict
          try { await graph("DELETE", `/users/${encodeURIComponent(room)}/events/${created.id}`); } catch {}
          return json(409, { error: "conflict" });
        }
      }

      // 4) branded confirmation from bookings@ (best-effort)
      try {
        const pretty = new Date(`${b.date}T00:00:00${TZ_OFFSET}`)
          .toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "Asia/Riyadh" });
        await graph("POST", `/users/${encodeURIComponent(SENDER_EMAIL)}/sendMail`, {
          body: {
            message: {
              subject: `Room booking confirmed — ${ROOM_LABEL.split("@")[0]} — ${pretty} (${b.start}–${b.end})`,
              body: { contentType: "HTML", content: confirmationHtml(b, room, ref, pretty) },
              toRecipients: [{ emailAddress: { address: b.email } }],
              ccRecipients: ADMIN_EMAIL && ADMIN_EMAIL !== b.email
                ? [{ emailAddress: { address: ADMIN_EMAIL } }] : [],
            },
            saveToSentItems: true,
          },
        });
      } catch (e) { console.error("sendMail failed (booking still valid):", e.status, e.data); }

      return json(200, { ok: true, ref });
    }

    return json(405, { error: "method not allowed" });
  } catch (err) {
    console.error("bookings error:", err.status, err.data || err.message);
    return json(500, { error: "server error" });
  }
};

function esc(s) {
  return String(s || "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
function confirmationHtml(b, room, ref, pretty) {
  return `
  <div style="font-family:Segoe UI,Arial,sans-serif;color:#152233;max-width:560px">
    <div style="background:#0b478d;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">
      <strong style="font-size:16px">${ORG_NAME}</strong><br>
      <span style="opacity:.85">Meeting Room Booking — confirmed</span>
    </div>
    <div style="border:1px solid #d8dee7;border-top:0;border-radius:0 0 10px 10px;padding:18px 20px">
      <p>Hi ${esc(b.name)}, your booking is confirmed.</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <tr><td style="color:#5a6675;padding:6px 0">Reference</td><td style="text-align:right;font-weight:700">${ref}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Room</td><td style="text-align:right;font-weight:700">${esc(room)}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Date</td><td style="text-align:right;font-weight:700">${pretty}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Time</td><td style="text-align:right;font-weight:700">${b.start} – ${b.end} (KSA)</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Attendees</td><td style="text-align:right;font-weight:700">${esc(b.attendees || "—")}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0;vertical-align:top">Purpose</td><td style="text-align:right;font-weight:700">${esc(b.purpose || "—")}</td></tr>
      </table>
      <p style="color:#5a6675;font-size:12px;margin-top:16px">To change or cancel, reply to this email.</p>
    </div>
  </div>`;
}
