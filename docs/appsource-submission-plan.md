# AppSource submission & cross-tenant test plan

Goal: get **Recipient Guard (Modern / New Outlook)** onto Microsoft AppSource as a
**free** add-in, and — as the pre-submission proof — install it on a **fresh
tenant** and confirm it works end to end there.

Owner tags: **[Fynn]** = only you can do it (Entra / Partner Center / billing /
legal). **[Claude]** = I can do it in the repo / Azure. **[Both]** = we pair.

The whole thing hinges on one prerequisite: **the Entra app must be
multi-tenant.** Today it's almost certainly single-tenant (the default), which is
why a new tenant can't use it yet. Fix that first and both the new-tenant test
*and* AppSource unblock.

---

## Phase 0 — Multi-tenant readiness (the prerequisite) ⭐

The add-in itself is fully hosted on Azure and needs no change. The work is in the
**Entra app registration** (client `38d26461-6bd7-4497-b83e-3a34baa9c154`).

- **[Fynn]** In Entra ID → App registrations → this app → **Authentication** →
  *Supported account types* → change to **"Accounts in any organizational
  directory (Any Microsoft Entra ID tenant – Multitenant)"**. (Add personal
  accounts too only if you want consumer Outlook.com later — not required for MVP.)
- **[Fynn]** Confirm the **SPA redirect** `brk-multihub://rgoutlookpoc0618.z33.web.core.windows.net`
  is present (it is, from production hosting). NAA uses the brokered redirect;
  no per-tenant redirect needed.
- **[Fynn]** API permissions: keep **delegated** `People.Read` + `User.Read`.
  These are **low-privilege, user-consentable** scopes — a normal user in the new
  tenant can consent without an admin (unless that tenant has disabled user
  consent, see Phase 1 gotchas). Do **not** add high-privilege scopes; they'd
  force admin consent and cross-tenant policy review.
- **[Fynn]** **Publisher verification** (MPN/Partner ID → "Verified publisher" blue
  badge on the consent screen). Not strictly required to *sideload*, but AppSource
  wants it and it removes the "unverified app" scare on the consent prompt. Start
  this early — it can take time.
- **[Fynn]** Set the **Publisher domain** on the app registration to a domain you
  own (helps consent UX + verification).

Exit check: from a browser signed into a *different* tenant, the consent screen
for this app appears (rather than "app not found in tenant").

---

## Phase 1 — New-tenant test (your stated next step)

Once the app is multi-tenant, prove it on a clean tenant. This is the best
pre-submission test — it exercises exactly what Microsoft's validators will hit.

- **[Fynn]** In the **new tenant**, sideload the production manifest:
  `manifest.azure.xml` via <https://aka.ms/olksideload> → My add-ins → Custom
  Addins → Add from File. (No admin-center deploy needed; per-mailbox sideload is
  instant.)
- **[Both]** Compose a test message and hit Send → confirm the **Smart Alert**
  fires (send-time blocking works cross-tenant — it's pure Office.js, no auth).
- **[Both]** Open the pane → **Turn on smart detection** → the **NAA consent
  prompt** should appear for the new tenant's user. Consent → `/me/people` loads
  → known contacts cached. This is the cross-tenant moment of truth.
- **[Both]** Re-test the full review flow on the new tenant: system dialog, Take
  action modal, whitelist, send, delay-with-confirmation.

**Gotchas to watch on the new tenant:**
- *"Approval required" / user can't consent* → that tenant has **user consent
  disabled** (Entra → Enterprise apps → Consent and permissions). Then a **tenant
  admin** must grant consent once (admin consent URL), after which all users work.
  Worth documenting for enterprise customers regardless.
- *"App not found in tenant"* → Phase 0 multi-tenant change hasn't propagated /
  wasn't saved.
- The **smart-detection** feature degrades gracefully if NAA is unavailable
  (older Outlook) — the external/same-name/same-username checks still work without
  Graph. Good to confirm the fallback too.

---

## Phase 2 — AppSource listing assets (I can draft most of this)

AppSource add-ins are **free-to-download only** (no in-store selling) — perfect for
the free MVP. Submission is via **Partner Center**.

**Required URLs — the #1 reason submissions fail is missing these:**
- **[Claude]** Draft a **Privacy Policy** and host it on the Azure site
  (`/privacy.html`). Our privacy story is genuinely strong and should be stated
  plainly: the add-in has **no backend**; it reads your *frequently-contacted
  people* via Microsoft Graph **People.Read** and caches a compact list in
  **Office roaming settings** (stays inside your Microsoft 365 mailbox); recipient
  analysis runs **locally** at send time; **no recipient data, email content, or
  contact data ever leaves your Microsoft 365 environment**. Must describe the
  app/service, not just a website.
- **[Claude]** Draft **Terms of Use / EULA** → host at `/terms.html`.
- **[Claude]** **Support URL** → a simple support/contact page (`/support.html`)
  or a mailto/support address. (Currently `SupportUrl` points at the task pane —
  change it to a real support page.)
- **[Both]** Add these URLs into the manifest and re-deploy.

**Listing content:**
- **[Both]** Add-in **name** ("Recipient Guard"), **short description**, **long
  description** (what it does, the AutoComplete-mistake problem it solves, privacy
  posture), **search keywords**, **category** (Productivity), supported
  **products/hosts** (Outlook), **languages** (en-US).
- **[Claude/Both]** **Screenshots** (1–5, 1366×768 recommended) — the review
  modal, the flagged recipient, the delay confirmation. I can produce clean
  mockups; real Outlook screenshots (like the ones you've been sending) are ideal.
- **[Fynn/Claude]** **Logos** — store logo (300×300) + the add-in icons already in
  the manifest. Current `icon.jpg` may need a proper square PNG at required sizes.

**Manifest for the store:**
- **[Claude]** Validate with the official validator (`npx office-addin-manifest
  validate manifest.azure.xml`) and fix anything it flags.
- **[Claude]** Ensure: all HTTPS (✓ Azure), no localhost (✓), valid `SupportUrl`,
  `AppDomains`, high-res icon, unique `<Id>`, sensible `<Version>`. Bump
  `<Version>` for each store resubmission.

---

## Phase 3 — Submit, validate, certify

- **[Fynn]** Create/confirm a **Partner Center** account and enrol in the
  **Microsoft 365 and Copilot** program (required to publish M365 apps). This
  needs company/verification details.
- **[Fynn]** In Partner Center → **Marketplace offers → New offer → Office add-in**
  → fill listing, upload manifest, add the URLs + assets from Phase 2.
- **[Both]** Provide **validation/test notes** for Microsoft's reviewers: what the
  add-in does, that **smart detection needs a Microsoft 365 account with some
  frequently-contacted people**, how to trigger a flag (compose to a look-alike
  address), and that it will prompt for **People.Read** consent. Missing/weak test
  notes is a common rejection cause.
- **Validation:** automated policy checks, then a human validation team tests on
  all supported platforms — **~3–5 business days**. Expect at least one round of
  feedback; iterate.
- **[Fynn]** **Microsoft 365 Certification / Publisher Attestation** — the
  attestation (self-declared security/privacy questionnaire) is the light path and
  is worth completing; it builds trust and some admins filter for it. Full
  certification is heavier and not required for launch.

---

## Suggested order

1. **Phase 0** (Fynn: multi-tenant + start publisher verification) — unblocks all.
2. **Phase 1** (new-tenant test) — proves cross-tenant consent before Microsoft does.
3. **Phase 2** (I draft privacy/terms/support pages + validate manifest; we do
   listing copy + screenshots) — can run in parallel with Phase 1.
4. **Phase 3** (Fynn: Partner Center enrolment; submit; iterate on validation).

## What I can start on right now
- Draft `privacy.html`, `terms.html`, `support.html` and wire them into the
  manifest + deploy.
- Run the manifest validator and fix issues.
- Draft the listing copy (short/long description, keywords) and a screenshot shot-list.

## Blocked on Fynn (external, I can't do)
- Entra multi-tenant change + publisher verification (Phase 0).
- Partner Center account + Microsoft 365/Copilot program enrolment (Phase 3).
- Final legal sign-off on the privacy policy / terms wording.
