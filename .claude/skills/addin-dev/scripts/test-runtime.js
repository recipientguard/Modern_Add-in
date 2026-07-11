// Offline test of the send-event analysis pipeline.
// Loads src/sendTestRuntime.js with a stubbed Office object and exercises the
// real onMessageSendDiagnostic handler across scenarios — no Outlook needed.
//
// Run from anywhere:  node .claude/skills/addin-dev/scripts/test-runtime.js
const path = require("path");

// repo root = up four levels from .claude/skills/addin-dev/scripts/
const repoRoot = path.resolve(__dirname, "..", "..", "..", "..");
const runtimePath = path.join(repoRoot, "src", "sendTestRuntime.js");

var RECIPS = { to: [], cc: [], bcc: [] };

global.Office = {
  AsyncResultStatus: { Succeeded: "succeeded" },
  context: {
    mailbox: {
      userProfile: { emailAddress: "fynn@iteam.je" },
      item: {
        to:  { getAsync: function (cb) { cb({ status: "succeeded", value: RECIPS.to }); } },
        cc:  { getAsync: function (cb) { cb({ status: "succeeded", value: RECIPS.cc }); } },
        bcc: { getAsync: function (cb) { cb({ status: "succeeded", value: RECIPS.bcc }); } }
      }
    }
  },
  actions: { associate: function () {} }
  // note: Office.onReady is intentionally absent so the runtime's onReady
  // registration path is skipped in this offline harness.
};

require(runtimePath);
var handler = globalThis.onMessageSendDiagnostic;
if (typeof handler !== "function") {
  console.error("FAIL: onMessageSendDiagnostic was not registered on globalThis");
  process.exit(1);
}

function r(name, email) { return { displayName: name, emailAddress: email }; }

var failures = 0;
function run(label, recips, expectAllow, expectContains, expectNotContains) {
  RECIPS = { to: recips.to || [], cc: recips.cc || [], bcc: recips.bcc || [] };
  return new Promise(function (resolve) {
    handler({ completed: function (result) { resolve(result); } });
  }).then(function (result) {
    var ok = result.allowEvent === expectAllow;
    var msg = result.errorMessage || "";
    if (ok && expectContains && msg.indexOf(expectContains) === -1) ok = false;
    if (ok && expectNotContains && msg.indexOf(expectNotContains) !== -1) ok = false;
    if (!ok) failures++;
    console.log((ok ? "PASS " : "FAIL ") + label + "  (allowEvent=" + result.allowEvent + ")");
    if (!ok && msg) console.log("      message: " + msg.replace(/\n/g, " | "));
  });
}

Promise.resolve()
  .then(function () { return run("clean internal send -> allow", { to: [r("Alice", "alice@iteam.je")] }, true); })
  .then(function () { return run("wrong-person name+prefix clash -> block", { to: [r("Fynn Hodder", "fynn@iteam.je"), r("Fynn Hodder", "fynn@gmail.com")] }, false, "possible wrong recipient"); })
  .then(function () { return run("external only -> block", { to: [r("Bob", "bob@client.com")] }, false, "external recipient"); })
  .then(function () { return run("personal prefix across domains -> block (prefix clash)", { to: [r("Jon A", "jon.doe@acme.com")], cc: [r("Jon B", "jon.doe@acme-invoices.com")] }, false, "Same address prefix"); })
  // Generic mailbox (sales@) across two real vendors is a legit send, not a
  // wrong-recipient clash: it should NOT produce a "Same address prefix" risk
  // (it's still flagged as external, which is correct).
  .then(function () { return run("generic prefix (sales@) across vendors -> not a prefix clash", { to: [r("Sales A", "sales@acme.com")], cc: [r("Sales B", "sales@partner.com")] }, false, "external recipient", "Same address prefix"); })
  // Regression: a prefix/display-name equal to an Object.prototype key must not
  // throw and fail the guard open (fixed by Object.create(null) grouping maps).
  .then(function () { return run("prototype-key prefix (constructor) -> block", { to: [r("A", "constructor@a.com")], cc: [r("B", "constructor@b.com")] }, false, "Same address prefix"); })
  .then(function () { return run("no recipients -> allow", {}, true); })
  .then(function () {
    console.log("");
    if (failures === 0) { console.log("All analysis tests passed."); }
    else { console.log(failures + " test(s) FAILED."); process.exit(1); }
  });
