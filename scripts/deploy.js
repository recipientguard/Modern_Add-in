// One-command deploy to Azure Static Web Apps (addin.recipientguard.co.uk).
//   npm run deploy
//
// Steps: build (runtime + MSAL bundle) -> assemble dist/ -> deploy dist/ to the
// Static Web App production environment.
//
// Auth: the deployment token is read from `az` at run time and passed to the SWA
// CLI via the SWA_CLI_DEPLOYMENT_TOKEN env var — never stored in the repo, never
// printed, and not exposed in the process args. Run `az login` first if needed.
//
// History: previously deployed to an Azure Storage static website
// (rgoutlookpoc0618). Storage can't serve HTTPS on a custom domain, so we moved to
// Static Web Apps, which provides a free managed certificate for
// addin.recipientguard.co.uk.

const { execFileSync, execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const SWA_NAME = "recipientguard-addin";
const RESOURCE_GROUP = "recipientguard-rg";
const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");

// Files the manifest (and the pages it loads) actually need.
const SRC_FILES = [
  "taskpane.html", "taskpane.css", "taskpane.js",
  "review-dialog.html", "review-dialog.js",
  "send-test.html", "sendTestRuntime.js", "recipientGuardCore.js",
  "naa.bundle.js"
];

// Public legal/support pages served at the site root.
const SITE_FILES = ["privacy.html", "terms.html", "support.html"];

// Icons referenced by the manifest (16/32/80 ribbon, 64 IconUrl, 128 hi-res) plus
// the 300x300 AppSource store logo. Regenerate from assets/icon.svg: npm run icons
const ICON_FILES = [
  "icon-16.png", "icon-32.png", "icon-64.png",
  "icon-80.png", "icon-128.png", "icon-215.png", "icon-300.png"
];

const INDEX_HTML =
  "<!doctype html>\n<html lang=\"en\"><head><meta charset=\"utf-8\">" +
  "<title>Recipient Guard</title></head><body style=\"font-family:system-ui;max-width:40rem;margin:4rem auto;padding:0 1rem\">" +
  "<h1>Recipient Guard</h1><p>Outlook add-in host. This site serves the add-in's pages; " +
  "there is nothing to see here directly.</p>" +
  "<p><a href=\"/privacy.html\">Privacy</a> &middot; <a href=\"/terms.html\">Terms</a> " +
  "&middot; <a href=\"/support.html\">Support</a></p></body></html>\n";

console.log("1/3  building...");
execSync("npm run build", { cwd: root, stdio: "inherit" });

console.log("2/3  assembling dist/...");
fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(dist, "src"), { recursive: true });
fs.mkdirSync(path.join(dist, "assets"), { recursive: true });
SRC_FILES.forEach((f) => fs.copyFileSync(path.join(root, "src", f), path.join(dist, "src", f)));
SITE_FILES.forEach((f) => fs.copyFileSync(path.join(root, "site", f), path.join(dist, f)));
ICON_FILES.forEach((f) => fs.copyFileSync(path.join(root, "assets", f), path.join(dist, "assets", f)));
fs.writeFileSync(path.join(dist, "index.html"), INDEX_HTML);
// Cache-Control: no-cache for the whole site. Outlook caches the runtime JS and
// the pane hard; stale assets have cost us hours. Revalidation (304s) is cheap.
fs.copyFileSync(path.join(root, "staticwebapp.config.json"), path.join(dist, "staticwebapp.config.json"));

console.log("3/3  deploying to Static Web App...");
// shell:true so Windows resolves az.cmd.
const token = execFileSync(
  "az",
  ["staticwebapp", "secrets", "list", "--name", SWA_NAME, "--resource-group", RESOURCE_GROUP,
   "--query", "properties.apiKey", "-o", "tsv"],
  { shell: true, encoding: "utf8" }
).trim();

execSync("npx -y @azure/static-web-apps-cli deploy ./dist --env production", {
  cwd: root,
  stdio: "inherit",
  env: { ...process.env, SWA_CLI_DEPLOYMENT_TOKEN: token }
});

console.log("\nDeployed to https://addin.recipientguard.co.uk/");
