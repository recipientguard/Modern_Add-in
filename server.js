const fs = require("fs");
const https = require("https");
const path = require("path");

const root = __dirname;
const certDir = path.join(root, ".certs");
const pfxPath = path.join(certDir, "localhost.pfx");
const passphrase = process.env.RECIPIENT_GUARD_DEV_CERT_PASSWORD || "recipientguard";
const port = Number(process.env.PORT || 3000);

// Tee every request line to a fixed log file so tooling (the addin-dev skill)
// can read it without knowing the background process's stdout path.
const requestLogPath = path.join(root, "dev-requests.log");
function logRequest(line) {
  console.log(line);
  try { fs.appendFileSync(requestLogPath, line + "\n"); } catch (e) { /* non-fatal */ }
}

const contentTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".xml", "application/xml; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".png", "image/png"],
  [".ico", "image/x-icon"],
  [".svg", "image/svg+xml"]
]);

function assertCertificateExists() {
  if (fs.existsSync(pfxPath)) return;

  console.error("Missing local HTTPS certificate: " + pfxPath);
  console.error("For local sideloading, create a localhost PFX certificate first.");
  console.error("Safer option: deploy this POC to Azure HTTPS and update manifest.xml URLs.");
  process.exit(1);
}

function serveFile(req, res) {
  const url = new URL(req.url, "https://localhost:" + port);
  let pathname = decodeURIComponent(url.pathname);
  logRequest(new Date().toISOString() + "  " + req.method + " " + pathname + (url.search || ""));
  if (pathname === "/__log") {
    res.writeHead(204, { "Cache-Control": "no-store", "Access-Control-Allow-Origin": "*" });
    res.end();
    return;
  }
  if (pathname === "/") pathname = "/src/taskpane.html";

  const filePath = path.normalize(path.join(root, pathname));
  if (!filePath.startsWith(root)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    res.writeHead(200, {
      "Content-Type": contentTypes.get(path.extname(filePath).toLowerCase()) || "application/octet-stream",
      "Cache-Control": "no-store"
    });
    res.end(data);
  });
}

assertCertificateExists();

https.createServer({
  pfx: fs.readFileSync(pfxPath),
  passphrase
}, serveFile).listen(port, () => {
  console.log("Recipient Guard New Outlook POC running at https://localhost:" + port);
  console.log("Task pane manifest: " + path.join(root, "manifest.taskpane-only.local.xml"));
  console.log("Send diagnostic manifest: " + path.join(root, "manifest.local-send-diagnostic.xml"));
});
