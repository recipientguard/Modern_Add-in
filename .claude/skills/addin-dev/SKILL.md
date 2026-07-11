---
name: addin-dev
description: >-
  Local dev loop and send-event diagnostics for the Recipient Guard "New Outlook"
  Office.js add-in (the Recipient_Guard_NewOutlook_POC repo). Use this whenever
  working on that add-in: starting or checking the local HTTPS dev server,
  figuring out why the OnMessageSend Smart Alert hangs / never fires / shows the
  wrong message, checking which manifest version is actually live, interpreting
  the request+beacon log, running the offline analysis tests, or preparing a
  manifest to sideload. Trigger on mentions of the add-in dev server, "send test",
  Smart Alerts, OnMessageSend, event-based activation, the manifest, sideloading,
  the beacon log, or "why isn't the add-in dispatching/firing". Prefer this skill
  over ad-hoc commands so the hard-won Office.js gotchas (especially the
  Office.onReady registration requirement) are applied automatically.
---

# Recipient Guard — New Outlook add-in dev & diagnostics

This add-in is an **Office.js Smart Alerts** add-in that blocks a send when it
detects a likely wrong recipient. Its runtime behaviour on new Outlook / Outlook
on the web is full of non-obvious traps, and the debug loop is slow (Outlook
caches aggressively, admin-center deployment takes many hours). This skill folds
the whole loop into a few reliable steps and encodes the gotchas we learned the
hard way.

Repo root: `C:\Users\FynnHodder\source\repos\Recipient_Guard_NewOutlook_POC`
(paths below are relative to it).

**Before deep debugging, read `references/office-js-gotchas.md`.** It holds the
root causes we already found — most send-event problems match one of them, so
checking it first saves hours.

## 1. Ensure the dev server is running

The add-in's manifests point at `https://localhost:3000`, served by `server.js`.
It stops whenever the laptop sleeps, so check and (re)start it first.

```bash
# Is it up?
curl -sk -o /dev/null -w "%{http_code}\n" https://localhost:3000/src/taskpane.html
```

- `200` → already running, continue.
- `000` / anything else → start it in the background:

```bash
node "C:/Users/FynnHodder/source/repos/Recipient_Guard_NewOutlook_POC/server.js"
```

Run it as a background process so it keeps serving while the user tests in
Outlook. It logs every request to `dev-requests.log` (in the repo root) as well
as stdout — read that file for diagnostics (next step). If start fails with a
missing-cert error, the dev cert at `.certs/localhost.pfx` is absent; regenerate
it with `scripts/create-dev-cert.ps1`.

## 2. Read and interpret the log

Every request Outlook makes to the local server is appended to
`dev-requests.log`. Two things matter:

**a) Which manifest is live** — from the `?v=` on the `send-test.html` request.
Each manifest tags its URLs with a distinct `?v=` (e.g. `?v=v3`), so the query
string tells you which deployed/sideloaded manifest actually loaded. If the user
"updated" the manifest but the log still shows the old `?v=`, the new one has not
propagated — see the deployment notes in the reference.

```bash
grep -E "send-test.html\?v=" dev-requests.log | tail -3
```

**b) How far the send handler got** — the runtime can phone progress home via
`beacon()` calls that hit the `/__log?stage=...` endpoint. When beacons are
enabled, the healthy trail is:

```
script-loaded;office=object;actions=yes   runtime loaded, Office ready
associate-ok                               top-level associate ran
onready-fired                              Office.onReady fired
associate-in-onready-ok                    re-registered here (the binding that matters)
handler-invoked                            Outlook dispatched the event
internal-domain;<domain>                   read the mailbox domain
recipients-read;n=<count>                  read To/Cc/Bcc
completing;allow=<true|false>              returned a decision
```

```bash
grep "stage=" dev-requests.log | tail -15
```

Interpret by where it **stops**:
- Stops after `associate-ok` (no `handler-invoked`) → **registered but never
  dispatched.** This is the classic new-Outlook/OWA failure — the handler must be
  associated inside `Office.onReady`. See the reference.
- No `stage=` lines at all, but you *do* see `send-test.html` fetched → the HTML
  runtime loaded but the runtime script isn't beaconing (beacons may be stripped;
  re-add them per the reference to trace further).
- Reaches `handler-invoked` but not `completing` → stuck in the recipient read;
  the safety timeout should still fail the send open after ~3s.
- Reaches `completing;allow=false` → working: it blocked with a real message.

Beacons are development-only and are normally stripped from `sendTestRuntime.js`.
To re-enable them for a debugging session, follow "Re-enabling beacons" in the
reference — the `/__log` endpoint in `server.js` is already there to receive them.

## 3. Editing the analysis engine: build first

The files the manifests load — `src/sendTestRuntime.js` and
`src/recipientGuardCore.js` — are **GENERATED** (see the banner at the top of
each). The single source of truth is `src/lib/engine.js` (shared analysis
engine) plus `src/lib/sendRuntime.part.js` / `src/lib/taskpaneCore.part.js`
(per-consumer glue). Never edit the generated files directly — the next build
would silently overwrite the change.

```bash
npm run build   # regenerates both bundles (scripts/build.js, zero deps)
```

Edit in `src/lib/`, run the build, then run the tests (next step). The build
outputs to the exact paths the live manifest already points at, so JS changes
need no manifest re-upload — just a browser hard-refresh.

## 4. Run the offline analysis tests

The recipient-analysis logic (external / same-display-name /
same-localpart-different-domain, with condensing) can be tested without Outlook.
This is the fast way to confirm a logic change before the slow Outlook round-trip.

```bash
npm test    # = node .claude/skills/addin-dev/scripts/test-runtime.js
```

It stubs `Office`, drives the real `onMessageSendDiagnostic` handler (from the
GENERATED `src/sendTestRuntime.js`, so build first) across the scenarios, and
prints PASS/FAIL with a non-zero exit on failure. If you change the engine, run
`npm run build && npm test`. If you add a rule or change messaging, add a
matching case to the harness.

## Typical flows

**"Is the add-in working / what's its state?"** → step 1 (ensure server up), then
step 2a (which manifest is live) + 2b (last beacon trail), then report plainly,
e.g. *"Server up; manifest v3 live; last send: handler-invoked → completing
allow=false (blocked, external recipient)."* Finish with step 4 so you can also
say the offline logic passes.

**"The send hangs / doesn't fire on new Outlook."** → read
`references/office-js-gotchas.md` first (the `Office.onReady` gotcha is the usual
culprit), then use the beacon trail (step 2b) to locate exactly where it stops.

**"I changed the analysis rules."** → edit `src/lib/`, then step 3 (build) and
step 4 (offline tests) before asking the
user to test in Outlook.

**"I need to test a manifest change in Outlook."** → see the deployment section of
the reference: sideload via `aka.ms/olksideload` (fast, per-mailbox) rather than
the admin center (slow), and give a stuck manifest a fresh `<Id>` + DisplayName to
escape the cache.
