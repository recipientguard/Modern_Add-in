# AppSource listing — copy, assets & shot-list

Everything to paste into the Partner Center offer listing for **Recipient Guard**.
Publisher: **RecipientGuard Ltd**. Free add-in, no in-app purchases.

---

## Names & identity

| Field | Value |
|---|---|
| **Offer / add-in name** | Recipient Guard |
| **Publisher name** | RecipientGuard Ltd |
| **Category** | Productivity (secondary: Communication) |
| **Products** | Outlook |
| **Language** | English (United States) |
| **Price** | Free |

---

## Short description
*(the one-liner under the title — keep it ~100 chars)*

> Warns you before you send to the wrong recipient picked from AutoComplete.

**Alternates:**
> Catches wrong recipients before you hit send — right inside Outlook.
> Stops the "wrong Dave" email. Checks your recipients at send time.

---

## Long description

> **Recipient Guard checks who you're actually emailing — before the message leaves.**
>
> AutoComplete is the single most common cause of misdirected email. You type "Dave",
> Outlook offers a Dave, and it's the wrong one. Recipient Guard pauses that send and
> shows you exactly who the message is going to.
>
> **What it catches**
> - **The wrong pick from AutoComplete** — you're emailing an address you don't normally
>   use for that person, even when it's the only recipient on the message.
> - **Look-alike recipients** — two people sharing a display name, or the same username
>   on different domains (john.doe@acme.com vs john.doe@acme-invoices.com).
> - **External recipients** — anyone outside your organisation.
>
> **How it works**
> When you hit Send, Recipient Guard reviews the recipients and — only if something
> looks off — pauses and shows you the full list, with the questionable one flagged and
> the reason why. You can send anyway, open a full review, or go back and fix it.
> Nothing to configure; it works from the moment it's installed.
>
> **Built for the way you actually send**
> - **Review before sending** — see every recipient, with the risky one highlighted.
> - **Delay 1 minute & send** — a cooling-off window with a cancel button, for when you
>   want a second look.
> - **Don't warn about this address** — whitelist an address you email deliberately, and
>   it stops asking.
> - **Smart detection** — optionally learns who you usually email (from your Microsoft 365
>   frequent contacts) so it can spot a single wrong recipient. One click to enable.
>
> **Privacy by design — there is no Recipient Guard server**
> Recipient Guard has no backend of its own. All recipient analysis runs locally inside
> Outlook. The only data it stores — your frequent-contact list and your whitelist — is
> kept inside your own Microsoft 365 mailbox using Office roaming settings. No recipient
> data, contact data, or message content is ever sent to us or to any third party. Message
> bodies and attachments are never read. See the privacy policy for detail.
>
> Free to install and use.

---

## Search keywords
*(Partner Center allows a handful — these are ranked)*

1. recipient
2. wrong recipient
3. misdirected email
4. autocomplete
5. email security
6. data loss prevention
7. send confirmation
8. email mistake

---

## Required URLs

| Field | URL |
|---|---|
| Privacy policy | https://addin.recipientguard.co.uk/privacy.html |
| Terms of use / EULA | https://addin.recipientguard.co.uk/terms.html |
| Support | https://addin.recipientguard.co.uk/support.html |
| Website / Home | https://recipientguard.co.uk/ |

---

## Logos & icons

Generated from `assets/icon.svg` — regenerate with `npm run icons`.

| Use | File | Size |
|---|---|---|
| **Store logo (AppSource)** | `assets/icon-300.png` | 300×300 |
| Manifest `IconUrl` | `assets/icon-64.png` | 64×64 |
| Manifest `HighResolutionIconUrl` | `assets/icon-128.png` | 128×128 |
| Ribbon icons | `icon-16/32/80.png` | 16/32/80 |

All live at `https://addin.recipientguard.co.uk/assets/<file>`.

---

## Screenshot shot-list

Partner Center takes **1–5** screenshots, **1366×768 PNG**. Capture in **Outlook on the
web** (cleanest chrome), light theme, on a tidy demo mailbox. Order matters — the first
is the hero.

Use plausible-but-fake names (e.g. *John Doe*, `john.doe@gmail.com` vs
`john.doe@acme.com`). **Don't show real customer addresses.**

1. **HERO — the send-time warning.** The grey "Recipient Guard paused this send" dialog
   over a compose window, showing the recipient list with one flagged
   ("You don't usually email this address") and the **Send anyway / Take action /
   Don't send** buttons. This is the whole product in one image.
2. **The review dialog.** The centered "Review before sending" modal — the combined
   recipient list, the amber-flagged recipient with its reason, and
   *"Don't warn about this address"*, plus Send / Delay 1 min & send / Cancel.
3. **Delay & send confirmation.** The "Sending in 60 seconds" view listing all
   recipients with the **Cancel send** button. Nice differentiator, shows the
   cooling-off idea.
4. **The task pane.** Recipient Guard pane with "Current recipients" and
   **Smart detection on — N known contacts loaded**. Shows it's configurable/live.
   *(Blur or fake the contact list — don't ship real addresses.)*
5. *(optional)* **Look-alike catch.** A compose with two similar recipients
   (`john.doe@acme.com` + `john.doe@acme-invoices.com`) flagged as
   "Same username as another recipient" — shows it works with no history at all.

**Tips:** full-window captures (not cropped mid-dialog), no notification toasts, no
personal mail visible in the list pane, and consistent zoom across all shots.

---

## Validation / test notes for Microsoft's reviewers

Paste something like this into the notes field — weak test notes are a common
rejection cause:

> Recipient Guard is a free Outlook add-in that warns the user at send time when a
> recipient looks like it may have been chosen incorrectly. No account, licence key, or
> sign-up is required — install and it works.
>
> **To see the core behaviour (no setup needed):**
> 1. Compose a message to any address outside the tester's own organisation, or to two
>    look-alike recipients (e.g. `john.doe@acme.com` and `john.doe@acme-invoices.com`).
> 2. Select **Send**. Recipient Guard pauses the send and shows a dialog listing the
>    recipients with the reason each was flagged.
> 3. Choose **Send anyway** to send as-is, **Don't send** to return to the message, or
>    **Take action** to open the full review pane/dialog (whitelist, send, or delay & send).
>
> **To test the optional "Smart detection" feature:**
> 1. Open the Recipient Guard task pane from the compose ribbon.
> 2. Select **Turn on smart detection**. Microsoft will prompt for consent to
>    **People.Read** and **User.Read** (delegated). Accept.
> 3. The add-in reads the signed-in user's frequently-contacted people via Microsoft
>    Graph to recognise when a recipient differs from the address normally used for that
>    person. This requires a Microsoft 365 account with some existing contact history.
>
> **Privacy:** the add-in has no backend service. Analysis runs client-side; the frequent
> contact list and whitelist are cached in Office roaming settings within the user's own
> mailbox. No data is transmitted to the publisher or any third party.
>
> **Platforms:** Outlook on the web and new Outlook on Windows (Mailbox requirement set
> 1.15). Support: support@recipientguard.co.uk

---

## Open product questions

- **"Smart detection (beta)"** — the pane still labels it *(beta)*. It's live-verified and
  works; decide whether to drop "(beta)" before the listing goes live (it reads as
  unfinished to a reviewer, but is honest). Currently left as-is.
- **Apex site** — `recipientguard.co.uk` is used as the Home page URL but isn't serving a
  site yet. A one-page marketing site there would tidy the listing.
