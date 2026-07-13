// One-command deploy to the Azure Storage static website.
//   npm run deploy
//
// Steps: build (runtime + MSAL bundle) -> assemble dist/ -> CLEAN MIRROR to the
// $web container (deletes stale blobs, uploads current, Cache-Control:no-cache).
//
// Auth: uses `az --auth-mode key` (auto-retrieves the account key; needs the
// listKeys permission your Azure role already has). Run `az login` first if
// needed. On Windows the `$web` container name is passed literally via cmd.

const { execFileSync, execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const ACCOUNT = "rgoutlookpoc0618";
const CONTAINER = "$web";
const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");

// Files the manifest (and the pages it loads) actually need.
const SRC_FILES = [
  "taskpane.html", "taskpane.css", "taskpane.js",
  "send-test.html", "sendTestRuntime.js", "recipientGuardCore.js",
  "naa.bundle.js"
];

const INDEX_HTML =
  "<!doctype html>\n<html lang=\"en\"><head><meta charset=\"utf-8\">" +
  "<title>Recipient Guard</title></head><body style=\"font-family:system-ui;max-width:40rem;margin:4rem auto;padding:0 1rem\">" +
  "<h1>Recipient Guard</h1><p>Outlook add-in host. This site serves the add-in's pages; " +
  "there is nothing to see here directly.</p></body></html>\n";

function az(args) {
  // shell:true so Windows resolves az.cmd; on Windows the shell is cmd, where
  // "$web" is a literal (not expanded).
  execFileSync("az", args, { stdio: "inherit", shell: true });
}

console.log("1/4  building...");
execSync("npm run build", { cwd: root, stdio: "inherit" });

console.log("2/4  assembling dist/...");
fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(dist, "src"), { recursive: true });
fs.mkdirSync(path.join(dist, "assets"), { recursive: true });
SRC_FILES.forEach((f) => fs.copyFileSync(path.join(root, "src", f), path.join(dist, "src", f)));
fs.copyFileSync(path.join(root, "assets", "icon.jpg"), path.join(dist, "assets", "icon.jpg"));
fs.writeFileSync(path.join(dist, "index.html"), INDEX_HTML);

console.log("3/4  clearing stale blobs in " + CONTAINER + "...");
az(["storage", "blob", "delete-batch", "--account-name", ACCOUNT, "--auth-mode", "key",
    "-s", CONTAINER, "--output", "none"]);

console.log("4/4  uploading current build...");
az(["storage", "blob", "upload-batch", "--account-name", ACCOUNT, "--auth-mode", "key",
    "-d", CONTAINER, "-s", dist, "--overwrite", "--content-cache-control", "no-cache",
    "--output", "none"]);

console.log("\nDeployed to https://" + ACCOUNT + ".z33.web.core.windows.net/");
