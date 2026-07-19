// Deploy the public marketing one-pager to its own Static Web App
// (recipientguard.co.uk).  npm run deploy:site
//
// Deliberately a SEPARATE Static Web App from the add-in host
// (addin.recipientguard.co.uk) so marketing tweaks never redeploy the add-in.
//
// The legal/support pages are deployed here too, from the SAME `site/` source the
// add-in host uses — one source of truth, so the two copies can't drift.
//
// Auth: deployment token read from `az` at run time via SWA_CLI_DEPLOYMENT_TOKEN —
// never stored or printed.

const { execFileSync, execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const SWA_NAME = "recipientguard-site";
const RESOURCE_GROUP = "rg-recipient-guard-prd";
const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist-site");

const SITE_FILES = ["privacy.html", "terms.html", "support.html"];
const ICON_FILES = ["icon-64.png", "icon-128.png", "icon-300.png"];

console.log("1/2  assembling dist-site/...");
fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(dist, "assets"), { recursive: true });
fs.copyFileSync(path.join(root, "www", "index.html"), path.join(dist, "index.html"));
SITE_FILES.forEach((f) => fs.copyFileSync(path.join(root, "site", f), path.join(dist, f)));
ICON_FILES.forEach((f) => fs.copyFileSync(path.join(root, "assets", f), path.join(dist, "assets", f)));
fs.copyFileSync(path.join(root, "staticwebapp.config.json"), path.join(dist, "staticwebapp.config.json"));

console.log("2/2  deploying to Static Web App...");
const token = execFileSync(
  "az",
  ["staticwebapp", "secrets", "list", "--name", SWA_NAME, "--resource-group", RESOURCE_GROUP,
   "--query", "properties.apiKey", "-o", "tsv"],
  { shell: true, encoding: "utf8" }
).trim();

execSync("npx -y @azure/static-web-apps-cli deploy ./dist-site --env production", {
  cwd: root,
  stdio: "inherit",
  env: { ...process.env, SWA_CLI_DEPLOYMENT_TOKEN: token }
});

console.log("\nDeployed the site to https://recipientguard.co.uk/");
