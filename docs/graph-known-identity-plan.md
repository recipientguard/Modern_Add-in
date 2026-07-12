# Plan: history-aware detection via Microsoft Graph (known-identity store)

Status: **planned** (not started). Written 2026-07-11. Execute in a focused session.

## Goal

Catch the **single wrong recipient** — the most common AutoComplete mistake — which the
current add-in misses. Example: composing to `fynn.hodder@onecollab.co.uk` when the user
normally reaches `fynn.hodder@iteam.je`. Today the add-in only compares recipients *against
each other within one email* (external + same-name/same-prefix among the current recipients),
so a lone wrong recipient just shows as "external".

Fix: give the add-in memory of **who the user usually emails**, and flag a recipient whose
name/prefix matches a known contact at a *different* address. This replicates the Classic
add-in's `KnownIdentityStore` (which read the Outlook AutoComplete cache + Contacts — sources
Office.js cannot touch).

## Approach (backend-free — confirmed by research)

- **Microsoft Graph via Nested App Authentication (NAA)** — MSAL.js
  `createNestablePublicClientApplication` gets a Graph token for the signed-in user **without a
  backend**. No server to build/host; the free-MVP posture holds.
- **Graph `/me/people`** (scope `People.Read`) returns the user's relevant/frequently-contacted
  people, ranked — the modern analog of the AutoComplete cache. This is the known-identity source.
- **`RoamingSettings`** (per-mailbox, stored server-side, loaded at Office init) is the shared
  cache readable from the task pane **and** the event runtimes — this is what fixes the
  cross-runtime sharing problem that broke the earlier localStorage approach.
- **NAA `acquireTokenSilent` works inside event runtimes** (`OnNewMessageCompose`,
  `OnMessageSend`). Interactive consent (`acquireTokenPopup`) **cannot** run in an event (no UI),
  so first-run consent must happen in the **task pane**.

## Data flow

1. **First run (task pane):** user clicks "Turn on smart detection" → NAA interactive consent for
   `People.Read` → fetch `/me/people` → write a compact known-identity list to `RoamingSettings`.
2. **On compose (`OnNewMessageCompose` event):** NAA `acquireTokenSilent` → refresh the
   known-identity list into `RoamingSettings` **if stale** (TTL, e.g. 24h). Runs well before send.
3. **On send (`OnMessageSend` event):** read the known-identity list from `RoamingSettings`
   (fast, local) → compare each recipient → add "known alternative" risks → PromptUser dialog.
   **No Graph call on the send path** (protects the ~5s budget).
4. **Degrade gracefully:** no consent yet, or NAA unsupported (`isSetSupported("NestedAppAuth",
   "1.1")` false) → fall back to today's within-message + external checks. Never block on auth.

## New detection rules (extend the engine)

`computeRisks(recipients, internalDomain, knownIdentities)` gains two rules, mirroring the
Classic `HasSameDisplayNameDifferentMailbox` / `HasSameLocalPartDifferentDomain` known-identity
branches:

- **`known_display_name`** — a recipient's `normalizedName` matches a known identity whose email
  differs → *"You usually reach 'Fynn Hodder' at fynn.hodder@iteam.je."*
- **`known_localpart`** — a recipient's `localPart` matches a known identity's localpart at a
  *different* domain → *"You usually email fynn.hodder@iteam.je."*

Both are **strong** signals (feed the existing `condense`/`isStrong`). Keep the generic-localpart
gating. These fire for a **single recipient**, which is the whole point.

Known-identity record (compact, to fit RoamingSettings size limit ~32KB — cap ~150 people):
`{ n: normalizedName, l: localPart, d: domain, e: email, name: displayName }`.

## Prerequisite — Entra app registration (Fynn, Azure portal, ~10 min)

1. portal.azure.com → **App registrations** → **New registration**.
2. **Name:** `Recipient Guard Add-in`. **Supported account types:** multitenant + personal MSAs
   (broadest; or single-tenant if only ever iteam.je).
3. **Redirect URI:** platform **Single-page application (SPA)**, URI `brk-multihub://localhost:3000`
   (origin only, no path). Add `brk-multihub://<azure-static-host>` for production later.
4. **Register** → copy the **Application (client) ID** (goes in the MSAL config).
5. **API permissions** → Add → Microsoft Graph → **Delegated** → `People.Read` (+ `User.Read`) →
   Grant.

## Phased implementation

- **A1 — Prove Graph (de-risk auth first).** Add `@azure/msal-browser`; init NAA in the task pane;
  a button that acquires a token and lists `/me/people` in the pane. Verify: interactive consent,
  then silent. **Also verify the round-trip:** write a value to `RoamingSettings` in the pane,
  read it back in the `OnMessageSend` event (confirms the cross-runtime cache before we build on
  it). This is the make-or-break step.
- **A2 — Known-identity store.** Fetch `/me/people` → compact list → `RoamingSettings` with a TTL;
  add the `OnNewMessageCompose` prefetch handler.
- **A3 — Detection.** Add `known_display_name` / `known_localpart` to `src/lib/engine.js`; pass the
  known list into `computeRisks`; wire the send runtime to load it from `RoamingSettings`; add the
  "you usually email…" messaging.
- **A4 — Test + polish.** Consent UX ("Turn on smart detection"), graceful degrade, new offline
  harness cases, live test in OWA.

## Build implication — we will need a real bundler

MSAL.js is an **npm dependency**, and the current zero-dependency concat build
(`scripts/build.js`) cannot bundle `node_modules` imports. This feature is the trigger to adopt a
proper bundler (recommend **esbuild** — tiny, fast, zero-config): it bundles `src/lib/engine.js` +
MSAL + the per-consumer glue into the runtime/taskpane files the manifests load. This supersedes
`scripts/build.js`. Plan for this in A1 (you can't add MSAL without it). Keep output ES2016-safe
where the send runtime is concerned (esbuild `--target=es2016`).

## AppSource / consent implications

- The `People.Read` Graph scope must be declared for the listing; Microsoft 365 certification
  reviews requested permissions.
- Need a **privacy policy** covering use of contact/relevance data (read-only, never leaves the
  client, only used to compare on-device).
- Users see a one-time consent prompt. Frame it clearly in the pane so it isn't a surprise.

## Open questions to close during A1

1. Confirm `RoamingSettings` written by the pane/compose-event is reliably readable in the
   `OnMessageSend` runtime on new Outlook + OWA (expected yes; verify empirically — this is the
   architecture's load-bearing assumption).
2. `OfficeRuntime.auth` vs `Office.auth` for the login hint in events (OfficeRuntime.auth has
   broader event support).
3. RoamingSettings practical size cap on the target clients → finalize the known-list cap.

## References

- NAA enablement: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- SSO in Outlook events (NAA) sample: https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/outlook-event-sso-naa/
- Graph from an Outlook add-in: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/microsoft-graph
- Persist state / RoamingSettings: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings
