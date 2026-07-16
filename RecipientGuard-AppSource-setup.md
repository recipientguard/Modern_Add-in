# Recipient Guard — Terms, URLs & AppSource setup

Reference for **RecipientGuard Ltd** — England & Wales, company number **17168880**,
registered office 71-75 Shelton Street, London WC2H 9JQ.
Support: **support@recipientguard.co.uk**

> The **live web pages are the source of truth** — this file is a convenience copy.
> Wording changes go in the repo (`Recipient_Guard_NewOutlook_POC/site/*.html`) and are
> republished with `npm run deploy:site`.

---

## Live URLs (www is canonical)

| Page | URL |
|---|---|
| Home / marketing | https://www.recipientguard.co.uk |
| Privacy Policy | https://www.recipientguard.co.uk/privacy.html |
| Terms of Use | https://www.recipientguard.co.uk/terms.html |
| Support | https://www.recipientguard.co.uk/support.html |
| Add-in host (assets only, not public-facing) | https://addin.recipientguard.co.uk |

`recipientguard.co.uk` (apex) **301-redirects to www** via GoDaddy forwarding — it works,
it just doesn't serve anything itself. Use the **www** URLs in Partner Center and the app
registration so machines get a direct 200 rather than a redirect hop.

---

## Icons (generated — `npm run icons` from `assets/icon.svg`)

| Use | File | Size |
|---|---|---|
| **Entra app registration logo** | `assets/icon-215.png` | **215×215** (Entra requires this, <100 KB) |
| **AppSource store logo** | `assets/icon-300.png` | 300×300 |
| Manifest IconUrl / hi-res | `icon-64.png` / `icon-128.png` | 64 / 128 |
| Ribbon | `icon-16/32/80.png` | 16 / 32 / 80 |

Repo path: `C:\Users\FynnHodder\source\repos\Recipient_Guard_NewOutlook_POC\assets\`

---

## ✅ Done

- Entra app **re-registered in the RecipientGuard tenant** — client `7519a415-3e8b-4c8e-9599-740a658ae7a2`,
  multitenant, SPA redirect `brk-multihub://addin.recipientguard.co.uk`, delegated
  `People.Read` + `User.Read` (admin-consented). Old iteam app deleted.
- Add-in hosted on **https://addin.recipientguard.co.uk** (Azure Static Web Apps, free managed cert).
- Marketing site live on **https://www.recipientguard.co.uk**; apex forwards to it.
- Legal + support pages published.
- Manifest **v1.3.0.0**, validates clean.
- Legacy stack decommissioned (old Next.js site, API, Postgres, ACR, container env, old storage).
  Recovery: app code in the repos + infra shape in `recipientguard-rg.arm.json` (this folder).
- Partner Center account created under RecipientGuard Ltd — **verification pending**.

## ⏳ To do — [Fynn]

1. **App registration → Branding & properties:**
   - Upload logo → `assets/icon-215.png`
   - Terms of service URL → `https://www.recipientguard.co.uk/terms.html`
   - Privacy statement URL → `https://www.recipientguard.co.uk/privacy.html`
   - Home page URL → `https://www.recipientguard.co.uk`
   - Confirm **Publisher domain = recipientguard.co.uk** (needed for publisher verification)
2. **Partner Center → Programs** — enrol in **Microsoft 365 and Copilot** (required to publish
   Office add-ins, even free ones).
3. **When verification clears → Account settings → Identifiers** — take the **Partner One ID**
   of type **PartnerGlobal** (not a location ID).
4. **Publisher verification** — app registration → Branding & properties → enter that Partner One
   ID → **Verify and save**. Removes the "unverified" warning from the consent screen.
5. **Capture 5 screenshots** (1366×768 PNG, clean demo mailbox, **fake names only**) — shot-list
   is in the repo at `docs/appsource-listing.md`.
6. **Submit the offer** — Partner Center → Marketplace offers → New offer → Office add-in.
   Upload `manifest.azure.xml`, the 300×300 logo, the URLs above, and the listing copy +
   reviewer test notes from `docs/appsource-listing.md`.

### Gotchas already hit (don't rediscover these)
- Partner Center's Legal business profile won't enable **Update** until the phone is **E.164**
  (`+447797791932`, not `0044…`), and the contact email needs its **OTP verified**.
- Legal/address fields are read-only during enrolment — edit them on the **Verification Summary**
  page afterwards.
- App registrations **cannot be moved between tenants** — re-registering mints a new client ID.
- Don't migrate the GoDaddy nameservers: the zone carries live **M365 MX + verification/SPF TXT**.
  Breaking mail also strands the Partner Center OTP.

---

## Terms of Use (current wording)

**Recipient Guard — Terms of Use**
_Last updated: 15 July 2026_

These terms are between you and **RecipientGuard Ltd**, registered in England and Wales
(company number 17168880) ("we", "us"). They govern your use of the Recipient Guard add-in for
Microsoft Outlook (the "add-in"). By installing or using the add-in, you agree to them.

**What the add-in does.** Recipient Guard checks the recipients of a message at send time and
warns you when one looks like it may have been chosen incorrectly. It is an assistive tool. You
remain solely responsible for the recipients of every message you send.

**No warranty.** The add-in is provided "as is", without warranties of any kind, express or
implied. It may not detect every wrong recipient, and it may occasionally flag a recipient that
is correct. It does not guarantee that any message is sent only to intended recipients.

**Limitation of liability.** To the maximum extent permitted by law, RecipientGuard Ltd is not
liable for any indirect, incidental, or consequential damages, or for any misdirected email, data
disclosure, or loss arising from use of (or reliance on) the add-in. The add-in is a safeguard,
not a substitute for checking your recipients.

**Acceptable use.** Use the add-in only within Microsoft Outlook and in line with your
organisation's policies and applicable law. Do not attempt to reverse engineer, resell, or misuse
the add-in.

**Cost.** The add-in is currently provided free of charge. These terms may be updated if that
changes; continued use after an update constitutes acceptance.

**Changes.** These terms may be updated from time to time. The "last updated" date above reflects
the current version.

**Contact.** Questions about these terms: support@recipientguard.co.uk.

---

## ⚠️ `recipientguard-rg.arm.json` (this folder)

Infrastructure export of the decommissioned stack — keep it, it's how the API/site would be
re-provisioned. **It contains the old Postgres admin password in clear text**, so keep it out of
repos and tickets. If that password is reused anywhere else, rotate it.
