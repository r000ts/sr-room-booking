# SR Room Booking — Microsoft Graph version

Booking page backed by the SystemRapid M365 tenant. Bookings land on the **room's
Outlook calendar**, availability is read live from it, and clashes are rejected at the
moment of booking. Confirmations are sent from **bookings@systemrapid.com**.

```
booking-graph/
├─ index.html                     ← booking page (FCC–Almabani branded, one pilot room)
├─ netlify/functions/bookings.js  ← Graph backend (availability + race-proof booking + email)
├─ package.json                   ← no dependencies (uses built-in fetch)
├─ netlify.toml
└─ README.md
```

## Already wired in (no need to set these unless they change)
| Setting | Value |
|---|---|
| Tenant ID | `0ccd0789-fd2c-48ff-9f4f-4b2bf5af070a` |
| Client ID | `71901e39-aa4d-4f67-a78f-707576c15925` |
| Room mailbox | `Area3_HQ-BoardMeeting@systemrapid.com` |
| Sender mailbox | `bookings@systemrapid.com` |
| Admin CC | `bookings@systemrapid.com` (override with `ADMIN_EMAIL` env) |

## The ONE secret you must set
In Netlify → **Site settings → Environment variables → Add**:

```
CLIENT_SECRET = <the client secret Value Subair copied at app registration>
```

Optional overrides (env vars): `ROOM_EMAIL`, `SENDER_EMAIL`, `ADMIN_EMAIL`, `ORG_NAME`.
Never put the secret in the code or in the front end — it lives only in this env var.

## Deploy
**Git (recommended):** push this folder, import the repo in Netlify, publish dir `.`,
set `CLIENT_SECRET`, deploy, then trigger one redeploy so the function picks up the var.

**CLI:**
```bash
npm i -g netlify-cli
cd booking-graph
netlify env:set CLIENT_SECRET "....the secret value...."
netlify deploy --prod
```

## Prerequisite from Subair (Exchange side)
Run the calendar-processing + access-policy script (the one from chat) so the room
auto-accepts and the app is scoped to the booking mailboxes only. The app-based path
blocks regardless, but `AutoAccept / AllowConflicts:$false` also protects anyone who
books the room the normal Outlook way.

## How blocking works
1. **Pre-check** the room calendar for the requested window → instant 409 on a clash.
2. **Create** the event on the room calendar.
3. **Verify**: re-read the window; if two events overlap, the racers tie-break by
   `createdDateTime` then `id` — every loser deletes its own event and returns 409.
   Exactly one survives, so two simultaneous bookings can never both win.
4. **Confirm**: branded email from `bookings@systemrapid.com` to the booker (CC admin).

## Test checklist (pilot)
1. Open the site → it should list today's bookings for the Area 3 Board Room (empty at first).
2. Make a booking → confirmation screen shows a ref; check the email arrives from bookings@;
   check the event appears on the room's Outlook calendar.
3. Try to book the **same slot** again → should show "that slot was just taken."
4. Optional hard test: two browsers, same slot, hit Confirm together → only one succeeds.

## Adding more rooms later
1. Subair creates each room mailbox + runs the same `Set-CalendarProcessing` and adds it
   to the `SR Booking Rooms` security group (the access policy then covers it automatically).
2. In `index.html` CONFIG, add `{ name: "...", email: "...@systemrapid.com" }` to `rooms`.
That's the whole change — no backend edit needed.
