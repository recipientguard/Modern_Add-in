# Recipient Guard New Outlook POC

This is a thin proof of concept for porting Recipient Guard to New Outlook.

It is intentionally separate from the Outlook Classic VSTO project.

## What this proves

- Loads as an Outlook web add-in.
- Adds a compose ribbon command.
- Runs an `OnMessageSend` Smart Alert handler.
- Reads To/Cc/Bcc recipients with Office.js.
- Derives the internal domain from the signed-in mailbox address.
- Warns when external recipients are detected.

## Local Dev

The local server requires a trusted HTTPS certificate at:

```text
.certs\localhost.pfx
```

Start the HTTPS static server:

```powershell
npm start
```

The server listens on:

```text
https://localhost:3000
```

For fast task-pane development, sideload this manifest directly in Outlook:

```text
manifest.taskpane-only.local.xml
```

It uses a separate add-in ID and the display name `Recipient Guard Local Dev`, so it can remain installed alongside the Azure POC. The local server sends `Cache-Control: no-store`, allowing JavaScript, HTML, and CSS changes to appear after reopening the task pane without updating the manifest.

The local task pane manifest supports pinning. After opening it from Outlook's Apps menu, pin it in the task pane so it stays open while composing.

## Local Send Event Diagnostic

Outlook event-based activation does not behave like a normal sideloaded task pane. Microsoft requires event-based add-ins to be admin-deployed before events such as `OnMessageSend` auto-run.

Use this manifest only when deliberately testing the send-time hook:

```text
manifest.local-send-diagnostic.xml
```

Deploy it through Microsoft 365 admin center > Settings > Integrated apps > Upload custom apps, assign it to `Just me`, then wait for Outlook to receive the update. This diagnostic manifest uses the display name `Recipient Guard Send Test` and `SendMode="Block"`.

The current local send flow is deliberately two-step to avoid Outlook timing out the send event:

1. Open `Recipient Guard Send Test` from Outlook's Apps menu.
2. Click `Check recipients` so the task pane reads To/Cc/Bcc and caches the result.
3. Click Send.

The send event only reads the cached result and completes immediately. Internal-only messages are allowed. If external recipients are found, clicking Send shows a Smart Alerts dialog that starts with:

```text
Recipient Guard found 1 external recipient.
```

Do not use `manifest.local-send-diagnostic.xml` for everyday task-pane work. Keep using `manifest.taskpane-only.local.xml` for fast UI changes.

## Development Certificate

For local sideloading, Outlook needs HTTPS. The helper script below creates a localhost certificate and imports it into the current user's trusted root store:

```powershell
.\scripts\create-dev-cert.ps1
```

That trust-store change is useful for local development, but it is still a security setting. The safer alternative is to deploy this POC to Azure HTTPS and update `manifest.xml` to point to the Azure URL.

## Sideload Manifest

Use this manifest for Azure testing:

`	ext
manifest.azure.xml
` 

Do not use the project-root manifest.xml unless you are deliberately testing the local https://localhost:3000 server.

## Azure Shape

For Azure, host these static files behind HTTPS and replace every `https://localhost:3000` manifest URL with the Azure app URL.

Recommended first Azure target:

- Azure Static Web Apps for the frontend
- Azure Functions or App Service for future licensing, identity memory, whitelist, and policy APIs

## Current Limitations

This POC only checks external recipients. Wrong-mailbox / AutoComplete ambiguity detection will need cloud-backed identity memory or Microsoft Graph contact lookup, because New Outlook add-ins cannot inspect local Outlook AutoComplete state like VSTO can.

