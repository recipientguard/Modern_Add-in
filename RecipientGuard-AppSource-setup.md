# Recipient Guard — Terms, URLs & AppSource setup

Reference for RecipientGuard Ltd (England & Wales, company number **17168880**).
Support contact: **support@recipientguard.co.uk**.

> The **live web pages are the source of truth** — this file is a convenience copy.
> If you change the wording, change it on the hosted pages (repo:
> `Recipient_Guard_NewOutlook_POC/site/*.html`) and re-deploy.

---

## Live URLs

| Page | URL |
|---|---|
| Privacy Policy | https://addin.recipientguard.co.uk/privacy.html |
| Terms of Use | https://addin.recipientguard.co.uk/terms.html |
| Support | https://addin.recipientguard.co.uk/support.html |

Hosted on **Azure Static Web Apps** (`recipientguard-addin` in `recipientguard-rg`)
with a free managed SSL certificate on the custom domain
**`addin.recipientguard.co.uk`**. Deploy with `npm run deploy` from the repo.

---

## Migration to the RecipientGuard Ltd tenant — status

**Why it was needed:** publisher verification (Step 2) requires the app's **tenant**
and **publisher domain** to match the domain verified in Partner Center. The app was
in the **iteam** tenant; the publisher is **RecipientGuard Ltd**.

**Key facts:**
- App registrations **cannot be moved** between tenants — this was a
  **re-registration** with a **new client ID**.
  - **New client ID:** `7519a415-3e8b-4c8e-9599-740a658ae7a2` (RecipientGuard tenant)
  - Old (iteam, retired): `38d26461-6bd7-4497-b83e-3a34baa9c154`
- Azure **Storage static websites can't serve HTTPS on a custom domain** (needs
  CDN/Front Door + BYO cert for an apex). So hosting moved to **Azure Static Web
  Apps**, which gives a **free managed certificate** on the custom domain.

### Done ✅
1. ✅ `recipientguard.co.uk` moved into the new RecipientGuard M365 tenant.
2. ✅ Hosting decided + built: **Azure Static Web Apps** (`recipientguard-addin`,
   RG `recipientguard-rg`) on **`https://addin.recipientguard.co.uk`**, free SSL.
   CNAME `addin` → `nice-field-08a514303.7.azurestaticapps.net`.
3. ✅ New app registration created in the RecipientGuard tenant, multitenant
   (all tenants), SPA redirect `brk-multihub://addin.recipientguard.co.uk`.
4. ✅ Claude wired it up: `CLIENT_ID` updated, all manifest URLs / `AppDomains` /
   `SupportUrl` → `addin.recipientguard.co.uk`, manifest `Version` → **1.1.0.0**,
   `deploy.js` retargeted at the Static Web App. Deployed; manifest validates clean.

### Remaining
5. **[Fynn]** On the **new** app registration confirm:
   - **API permissions (delegated):** `People.Read`, `User.Read`
   - **Branding & properties:** Home page `https://recipientguard.co.uk/`,
     **Terms** + **Privacy** URLs (the `addin.recipientguard.co.uk` ones above),
     **Publisher domain = recipientguard.co.uk**
6. **[Fynn]** **Re-sideload** the updated `manifest.azure.xml` (remove the old
   add-in first to avoid a cached manifest) and **re-test** on the RecipientGuard
   tenant *and* a second tenant: consent prompt → contacts load → send-block +
   review flow.
7. **[Fynn]** Once confirmed, **delete the old iteam app registration** and
   decommission the old storage site (`rgoutlookpoc0618`).
8. Then → **Step 3** (Partner Center) → **Step 2** (publisher verification).

---

## Terms of Use (current wording)

**Recipient Guard — Terms of Use**
_Last updated: 15 July 2026_

These terms are between you and **RecipientGuard Ltd**, registered in England and
Wales (company number 17168880) ("we", "us"). They govern your use of the Recipient
Guard add-in for Microsoft Outlook (the "add-in"). By installing or using the
add-in, you agree to them.

**What the add-in does.** Recipient Guard checks the recipients of a message at
send time and warns you when one looks like it may have been chosen incorrectly. It
is an assistive tool. You remain solely responsible for the recipients of every
message you send.

**No warranty.** The add-in is provided "as is", without warranties of any kind,
express or implied. It may not detect every wrong recipient, and it may
occasionally flag a recipient that is correct. It does not guarantee that any
message is sent only to intended recipients.

**Limitation of liability.** To the maximum extent permitted by law, RecipientGuard
Ltd is not liable for any indirect, incidental, or consequential damages, or for
any misdirected email, data disclosure, or loss arising from use of (or reliance
on) the add-in. The add-in is a safeguard, not a substitute for checking your
recipients.

**Acceptable use.** Use the add-in only within Microsoft Outlook and in line with
your organisation's policies and applicable law. Do not attempt to reverse
engineer, resell, or misuse the add-in.

**Cost.** The add-in is currently provided free of charge. These terms may be
updated if that changes; continued use after an update constitutes acceptance.

**Changes.** These terms may be updated from time to time. The "last updated" date
above reflects the current version.

**Contact.** Questions about these terms: support@recipientguard.co.uk.

---

## Step 1 — put the URLs on the app registration ✅ (done on the iteam app)

This removes the *"publisher has not provided links to their terms"* line on the
sign-in consent screen.

> ✅ Done on the **old iteam app** (Terms + Privacy URLs saved, Home page set to
> `https://recipientguard.co.uk/`).
> ⚠️ **Will need doing again on the new RecipientGuard-tenant app** — these settings
> don't carry across tenants.

1. **entra.microsoft.com** → **App registrations** → **Recipient Guard Add-in**.
2. **Branding & properties**.
3. Set **Terms of service URL** = the Terms URL above.
4. Set **Privacy statement URL** = the Privacy URL above.
5. **Save**. (Also confirm **Publisher domain** is set — needed for Step 2.)

---

## Step 2 — Publisher verification (removes the "unverified" warning)

Verification replaces *"unverified / not published by Microsoft or your
organisation"* on the consent screen with your verified company name + a badge.
It's **free** and quick **once the prerequisites are in place** — but the main
prerequisite is a verified partner account, which is Step 3. **So do Step 3 first,
then come back here.**

**What you need:**
- A **Partner One ID** from a **Microsoft AI Cloud Partner Program** account (the
  old "MPN") that has completed verification. It must be the **Partner Global
  Account (PGA)** ID — in Partner Center: *Account settings → Identifiers →
  Microsoft Cloud Partner Program → the ID with type **PartnerGlobal*** (not a
  "location" ID).
- The **Entra tenant** that holds the app registration must be **associated with
  that PGA**.
- The app's **Publisher domain** must **match** the domain verified in Partner
  Center. ⚠️ **Important for us:** the app is currently registered in the **iteam**
  tenant. Decide whether the publisher is **iteam** or **RecipientGuard Ltd** and
  keep tenant + publisher domain + Partner account consistent. If RecipientGuard
  Ltd is the publisher, the cleanest setup is the app registration living in a
  RecipientGuard Ltd tenant with `recipientguard.co.uk` as the verified/publisher
  domain. Worth sorting before verifying — it avoids a domain-mismatch error.
- Your sign-in account needs the Entra role **Application Administrator** or **Cloud
  Application Administrator**, be a **Partner Center admin** (CPP Admin or Accounts
  Admin), and use **MFA**.

**Steps:**
1. Sign in (with MFA) to **entra.microsoft.com** → App registrations → the app →
   **Branding & properties**.
2. Confirm the **Publisher domain** is set and verified.
3. In the **Publisher verification** section, choose **Add MPN ID / Verify**, enter
   your **Partner One ID (PartnerGlobal)**, then **Verify and save**.
4. If it errors with a domain mismatch, align the app's publisher domain with the
   Partner Center verified email domain and retry.

Docs: <https://learn.microsoft.com/entra/identity-platform/mark-app-as-publisher-verified>

---

## Step 3 — Partner Center account + Microsoft 365 and Copilot program

This is what lets you publish to Microsoft Marketplace / AppSource, and it's what
produces the Partner One ID that Step 2 needs. **Do this first.**

**Before you start:** you need authority to sign legal agreements for RecipientGuard
Ltd, and the company's **legal business name, registered address, and primary
contact**. You must use a **work/school (Microsoft Entra) account — not a personal
Microsoft account.**

**Steps:**
1. Go to **partner.microsoft.com** and **sign up / create a Partner Center
   account** using a RecipientGuard Ltd work account. Enter the legal business
   details (name **RecipientGuard Ltd**, company number 17168880, registered
   address).
2. Complete Microsoft's **business verification** (they confirm the company and
   your authority — this is the part that can take a little time).
3. Enrol in the **Microsoft 365 and Copilot** program (required to publish Office
   add-ins, even free ones). Read and **Accept** the **Microsoft Publisher
   Agreement**.
4. Once enrolled, your partner profile shows the **Partner One ID** — use it for
   Step 2 (publisher verification).
5. Then create the offer: Partner Center → **Marketplace offers → New offer →
   Office add-in** → upload `manifest.azure.xml`, add the Privacy/Terms/Support
   URLs, listing copy, screenshots, and validator/test notes.

Docs:
- Open an Office account: <https://learn.microsoft.com/partner-center/marketplace-offers/open-a-developer-account>
- Submit to AppSource: <https://learn.microsoft.com/partner-center/marketplace-offers/submit-to-appsource-via-partner-center>

---

## Recommended order (updated)

0. **⏳ CURRENT — tenant migration** (see the section above): domain into the new
   tenant → decide hosting → re-register the app → new client ID to Claude →
   rebuild/redeploy → re-test → delete the old app. **Everything else waits on this.**
1. **Step 3** — create the Partner Center account under RecipientGuard Ltd + enrol
   in the Microsoft 365 and Copilot program (this is what gives you the **Partner
   One ID** that Step 2 needs — so it comes first).
2. **Step 2** — publisher verification using that Partner One ID (only works once
   the app lives in the RecipientGuard tenant with `recipientguard.co.uk` as the
   publisher domain — hence step 0).
3. **Step 1** — re-apply the Terms/Privacy/Home URLs on the **new** app registration.
4. Build the store listing (copy, screenshots, icon) and submit the offer.

## Open decisions

- **Hosting domain:** stay on the Azure blob URL, or move to `recipientguard.co.uk`?
  (Decide before re-registering — it sets the SPA redirect URI.)
- **Icon:** current `icon.jpg` passes validation, but a proper square **PNG** is
  better for the store listing.
- **Pane subtitle** still reads *"Proof of concept for New Outlook send-time
  recipient checks."* — should be depersonalised before the listing.
