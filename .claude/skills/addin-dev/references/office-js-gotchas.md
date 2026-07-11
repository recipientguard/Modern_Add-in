# Office.js Smart Alerts gotchas (Recipient Guard, learned the hard way)

Read this before debugging a send-event problem. Each section is a real trap that
cost hours; most `OnMessageSend` failures match one of them.

## Contents
1. The `Office.onReady` registration gotcha (the big one)
2. Platform runtime split (which file each client loads)
3. Manifest requirements for `OnMessageSend`
4. Deployment: sideload vs admin center, and busting a stuck cache
5. The beacon-diagnostic method (and how to re-enable it)
6. Why the analysis runs inline at send time

---

## 1. The `Office.onReady` registration gotcha ⭐

**Symptom:** On new Outlook / Outlook on the web, clicking Send shows Outlook's
*"<add-in> is taking longer than expected … is processing message"* dialog and
hangs. The runtime clearly loads and `Office.actions.associate(...)` runs without
error, but the handler is never actually called.

**Root cause:** On new Outlook and OWA, `Office.actions.associate` only binds the
handler when it is called **inside `Office.onReady()`**. A top-level call — which
is exactly what the `yo office` template generates — returns successfully but
Outlook never dispatches the event to it.

**Fix:** register in BOTH places. Top level covers classic Outlook; the
`Office.onReady` call is what makes new Outlook/OWA dispatch.

```js
function register() {
  try { Office.actions.associate("onMessageSendDiagnostic", onMessageSendDiagnostic); }
  catch (e) { /* not ready yet; onReady covers it */ }
}
register();                                   // classic Outlook
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(register);                   // new Outlook / OWA — the binding that matters
}
```

This was the fix for the multi-day hang. If a send hangs and the beacon trail
stops after `associate-ok` with no `handler-invoked`, this is almost certainly it.

## 2. Platform runtime split

For event-based activation each Outlook client loads a **different file**, per
Microsoft's docs:

- **Classic Outlook on Windows** runs the **JavaScript file** directly
  (`<Override type="javascript" resid="Commands.Script.Url">` → `sendTestRuntime.js`).
- **New Outlook on Windows, Outlook on the web, new Mac UI** run the **HTML file**
  (`<Runtime resid="Commands.Url">` → `send-test.html`), which loads `office.js`
  plus the runtime script via `<script>` tags.

Consequences:
- The runtime JS must be **self-contained** — on the JS-file path there is no HTML
  to pull in shared modules, so don't depend on another script being loaded.
- Classic caches the JS file in the Wef folder and may take up to 24h to refresh;
  OWA re-fetches from the server (respecting `Cache-Control: no-store`), so OWA is
  the fast client to iterate/debug on.

## 3. Manifest requirements for `OnMessageSend`

Match Microsoft's official Smart Alerts sample structure in the add-in-only XML
manifest:
- Requirement set **`DefaultMinVersion="1.15"`** in the VersionOverrides
  (`OnMessageSend` exists from 1.12, but 1.15 is what MS uses for Smart Alerts).
- The `<Runtime>` element must **not** carry a `lifetime` attribute in the XML
  manifest (that attribute belongs to the JSON unified manifest; in XML it does
  not belong on `<Runtime>`).
- `SendMode="SoftBlock"` (blocks until fixed, but sends if the add-in is
  unavailable) is the known-good value. `"Block"` is valid but harsher;
  `"PromptUser"` gives a "send anyway" override for a warn-not-block UX.
- The `LaunchEvent` `FunctionName` must exactly match the name passed to
  `Office.actions.associate`.

Note: none of these manifest details were the cause of the hang (we tried 1.12 vs
1.15, with/without `lifetime`, Block vs SoftBlock). They are correctness/parity
items; the actual dispatch fix was #1. But keep them aligned to avoid new issues.

## 4. Deployment: sideload vs admin center, and busting a stuck cache

**Admin center (admin.microsoft.com → Integrated apps) propagation is glacial**
for event-based add-ins — both *add* and *remove* observed to take 20+ hours, and
a deleted app kept showing in the ribbon. Do not rely on it for the dev loop.

**Use per-mailbox sideloading instead** — it propagates in seconds:
`https://aka.ms/olksideload` → **My add-ins** → **Custom Addins** →
**Add a custom add-in** → **Add from File** → pick the manifest.

**If a manifest is stuck** (Outlook keeps loading the old one — the log shows the
old `?v=`): give the manifest a **new `<Id>` (fresh GUID) and a new DisplayName**.
Outlook keys off `<Id>`, so re-uploading the same Id can collide with the cached
old one; a new Id is treated as a brand-new add-in with no cache history. That is
how the working `manifest.local-send-diagnostic-v3.xml` was made from v2.

**Confirming which manifest is live:** give each manifest a distinct `?v=` on its
URLs; the server request log then shows exactly which one loaded.

## 5. The beacon-diagnostic method

The event runtime is opaque — no F12/console access on desktop, and even in OWA
the event runtime's console does not surface in the compose window's DevTools. So
the runtime **beacons its progress to the local server**, which logs it.

Mechanism:
- `server.js` has a `/__log` endpoint that 204s and logs the query string, plus it
  tees all requests to `dev-requests.log`.
- The runtime calls a tiny `beacon(stage)` helper that does a fire-and-forget
  `fetch("https://localhost:3000/__log?stage=" + stage, {mode:"no-cors"})`.

**Re-enabling beacons** (they are stripped from `sendTestRuntime.js` once a path is
confirmed working): add the helper near the top of the runtime IIFE and sprinkle
`beacon("...")` calls at the key points (script load, after each associate, at
handler entry, after the recipient read, before `event.completed`). Then have the
user send from OWA and read `dev-requests.log`. See the git history around the
"register in Office.onReady" fix for the exact instrumentation that cracked it.

This beats guessing: it tells you precisely whether the script loaded, whether
Office was ready, whether the handler registered, whether Outlook dispatched, and
where it stopped.

## 6. Why the analysis runs inline at send time

An earlier design had the task pane compute the analysis and stash it in
`localStorage` for the send handler to read (a two-step "click Check recipients
first" flow). That is fragile: separate Outlook runtimes don't reliably share
`localStorage` (on classic the send event saw an empty cache). The current
runtime reads To/Cc/Bcc **directly at send time** via
`Office.context.mailbox.item.*.getAsync` and runs the analysis inline, then
completes. It has a fail-open safety timeout (~3s, under Outlook's 5s limit) so a
stalled read can never hang the send. Reading recipients is sub-second, so the
timing-budget worry that motivated the two-step design was overblown.
