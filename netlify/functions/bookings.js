// ============================================================
//  Netlify Function: /.netlify/functions/bookings   (Microsoft Graph)
//
//  Works entirely through the bookings@ mailbox (which Graph can resolve),
//  so it does NOT depend on Graph being able to look up the room as a /users
//  object. The room is reached by SMTP via free/busy + resource invitations.
//
//  GET  ?date=YYYY-MM-DD  -> { bookings:[{start,end,name}] }
//  POST { name,email,date,start,end,attendees,purpose }
//         -> 200 { ok, ref }  |  409 { error:"conflict" }
//
//  Blocking: bookings@ creates the meeting and invites the room as a resource.
//  Exchange's resource booking attendant (AutoAccept / AllowConflicts:$false)
//  accepts if free, declines if it clashes. We read the room's response:
//  declined -> 409 (and we delete our event). Pre-check via getSchedule gives
//  a fast 409 on obvious clashes before we ever create anything.
// ============================================================

const TENANT_ID = process.env.TENANT_ID || "0ccd0789-fd2c-48ff-9f4f-4b2bf5af070a";
const CLIENT_ID = process.env.CLIENT_ID || "71901e39-aa4d-4f67-a78f-707576c15925";
const CLIENT_SECRET = process.env.CLIENT_SECRET;            // env only
const ROOM_SMTP   = process.env.ROOM_SMTP   || "Area3_HQ-BoardMeeting@systemrapid.com";
const ROOM_LABEL  = process.env.ROOM_LABEL  || "Area 3 Board Room";
const SENDER_EMAIL = process.env.SENDER_EMAIL || "bookings@systemrapid.com";
const ADMIN_EMAIL  = process.env.ADMIN_EMAIL  || "bookings@systemrapid.com";
const ORG_NAME     = process.env.ORG_NAME     || "FCC\u2013Almabani Joint Venture";
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

let _tok = { value: null, exp: 0 };
async function getToken() {
  if (_tok.value && Date.now() < _tok.exp - 60000) return _tok.value;
  const body = new URLSearchParams({
    client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials",
  });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body,
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
  const r = await fetch(GRAPH + path, { method, headers, body: body ? JSON.stringify(body) : undefined });
  const text = await r.text();
  const data = text ? JSON.parse(text) : {};
  if (!r.ok) { const e = new Error(`graph ${method} ${path}: ${r.status}`); e.status = r.status; e.data = data; throw e; }
  return data;
}

const hm = (dt) => String(dt).slice(11, 16);
const overlaps = (aS, aE, bS, bE) => aS < bE && aE > bS;
const sleep = (ms) => new Promise((res) => setTimeout(res, ms));

// free/busy for the room over a window, via the bookings@ mailbox
async function getRoomSchedule(startLocal, endLocal) {
  const data = await graph("POST", `/users/${encodeURIComponent(SENDER_EMAIL)}/calendar/getSchedule`, {
    prefer: `outlook.timezone="${TZ_WINDOWS}"`,
    body: {
      schedules: [ROOM_SMTP],
      startTime: { dateTime: startLocal, timeZone: TZ_WINDOWS },
      endTime:   { dateTime: endLocal,   timeZone: TZ_WINDOWS },
      availabilityViewInterval: 30,
    },
  });
  const sched = (data.value && data.value[0]) || {};
  const items = (sched.scheduleItems || []).filter(
    (i) => i.status && i.status !== "free");
  return { items, availabilityView: sched.availabilityView || "" };
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return json(200, {});
  if (!CLIENT_SECRET) return json(500, { error: "server not configured (CLIENT_SECRET missing)" });

  try {
    // -------------------- availability --------------------
    if (event.httpMethod === "GET") {
      const date = (event.queryStringParameters || {}).date;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date || "")) return json(400, { error: "date required (YYYY-MM-DD)" });
      const { items } = await getRoomSchedule(`${date}T00:00:00`, `${date}T23:59:59`);
      const bookings = items.map((i) => ({
        start: hm(i.start.dateTime),
        end: hm(i.end.dateTime),
        name: (i.subject || "Booked").split("  [")[0].split(" \u2014 ")[0],
      }));
      return json(200, { bookings });
    }

    // -------------------- create booking --------------------
    if (event.httpMethod === "POST") {
      let b;
      try { b = JSON.parse(event.body || "{}"); } catch { return json(400, { error: "invalid JSON" }); }
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
      const pre = await getRoomSchedule(startLocal, endLocal);
      const preClash = pre.items.some((i) =>
        overlaps(i.start.dateTime, i.end.dateTime, startLocal, endLocal))
        || /[^0\s]/.test(pre.availabilityView);
      if (preClash) return json(409, { error: "conflict" });

      // 2) bookings@ creates the meeting, room invited as a resource
      const created = await graph("POST", `/users/${encodeURIComponent(SENDER_EMAIL)}/events`, {
        body: {
          subject: `${b.name} \u2014 ${b.purpose ? b.purpose : "Meeting"}  [${ref}]`,
          body: { contentType: "HTML", content:
            `Booked via the room booking portal.<br>Ref: ${ref}<br>`
            + `Booker: ${esc(b.name)} &lt;${esc(b.email)}&gt;<br>`
            + `Attendees: ${esc(b.attendees || "\u2014")}<br>Purpose: ${esc(b.purpose || "\u2014")}` },
          start: { dateTime: startLocal, timeZone: TZ_WINDOWS },
          end:   { dateTime: endLocal,   timeZone: TZ_WINDOWS },
          location: { displayName: ROOM_LABEL },
          attendees: [
            { type: "resource", emailAddress: { address: ROOM_SMTP, name: ROOM_LABEL } },
          ],
          transactionId: ref,
        },
      });

      // 3) read the room's response (auto-accept / auto-decline)
      let status = "none";
      for (let i = 0; i < 6; i++) {
        await sleep(800);
        let ev;
        try {
          ev = await graph("GET",
            `/users/${encodeURIComponent(SENDER_EMAIL)}/events/${created.id}?$select=attendees`);
        } catch { continue; }
        const room = (ev.attendees || []).find((a) =>
          (a.emailAddress?.address || "").toLowerCase() === ROOM_SMTP.toLowerCase());
        status = room?.status?.response || "none";
        if (status === "declined") {
          try { await graph("DELETE", `/users/${encodeURIComponent(SENDER_EMAIL)}/events/${created.id}`); } catch {}
          return json(409, { error: "conflict" });
        }
        if (status === "accepted" || status === "organizer") break;
      }

      // 3b) if still undecided, verify the room is actually busy now; if not, fail safe
      if (status !== "accepted" && status !== "organizer") {
        const check = await getRoomSchedule(startLocal, endLocal);
        const busy = check.items.some((i) =>
          overlaps(i.start.dateTime, i.end.dateTime, startLocal, endLocal))
          || /[^0\s]/.test(check.availabilityView);
        if (!busy) {
          try { await graph("DELETE", `/users/${encodeURIComponent(SENDER_EMAIL)}/events/${created.id}`); } catch {}
          return json(500, { error: "room did not confirm; please retry" });
        }
      }

      // 4) branded confirmation from bookings@ (best-effort)
      try {
        const pretty = new Date(`${b.date}T00:00:00${TZ_OFFSET}`)
          .toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "Asia/Riyadh" });
        await graph("POST", `/users/${encodeURIComponent(SENDER_EMAIL)}/sendMail`, {
          body: {
            message: {
              subject: `Room booking confirmed \u2014 ${ROOM_LABEL} \u2014 ${pretty} (${b.start}\u2013${b.end})`,
              body: { contentType: "HTML", content: confirmationHtml(b, ref, pretty) },
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
  return String(s || "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
function confirmationHtml(b, ref, pretty) {
  return `
  <div style="font-family:Segoe UI,Arial,sans-serif;color:#152233;max-width:560px">
    <div style="background:#0b478d;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">
      <strong style="font-size:16px">${ORG_NAME}</strong><br>
      <span style="opacity:.85">Meeting Room Booking \u2014 confirmed</span>
    </div>
    <div style="border:1px solid #d8dee7;border-top:0;border-radius:0 0 10px 10px;padding:18px 20px">
      <p>Hi ${esc(b.name)}, your booking is confirmed.</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <tr><td style="color:#5a6675;padding:6px 0">Reference</td><td style="text-align:right;font-weight:700">${ref}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Room</td><td style="text-align:right;font-weight:700">${esc(ROOM_LABEL)}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Date</td><td style="text-align:right;font-weight:700">${pretty}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Time</td><td style="text-align:right;font-weight:700">${b.start} \u2013 ${b.end} (KSA)</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0">Attendees</td><td style="text-align:right;font-weight:700">${esc(b.attendees || "\u2014")}</td></tr>
        <tr><td style="color:#5a6675;padding:6px 0;vertical-align:top">Purpose</td><td style="text-align:right;font-weight:700">${esc(b.purpose || "\u2014")}</td></tr>
      </table>
      <p style="color:#5a6675;font-size:12px;margin-top:16px">To change or cancel, reply to this email.</p>
    </div>
  </div>`;
}
