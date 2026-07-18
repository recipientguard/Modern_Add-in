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
var ROAMING = {}; // stubbed RoamingSettings store (known-identity cache)

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
    },
    roamingSettings: {
      get: function (key) { return ROAMING[key]; },
      set: function (key, value) { ROAMING[key] = value; },
      saveAsync: function (cb) { if (cb) cb({ status: "succeeded" }); }
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

// Build a compact known-identity record like engine.toKnownRecord would.
function knownRec(name, email) {
  var e = (email || "").trim().toLowerCase();
  var at = e.lastIndexOf("@");
  return {
    name: name,
    e: e,
    n: (name || "").toLowerCase().replace(/[^a-z0-9]/g, ""),
    l: at > 0 ? e.slice(0, at) : "",
    d: (at > -1 && at < e.length - 1) ? e.slice(at + 1) : ""
  };
}
function setKnown(records) {
  ROAMING["recipientGuard.knownIdentities.v1"] =
    records && records.length ? { at: Date.now(), people: records } : undefined;
}
function setWhitelist(emails) {
  ROAMING["recipientGuard.whitelist.v1"] =
    emails && emails.length ? { at: Date.now(), emails: emails, domains: [] } : undefined;
}
function setWhitelistDomains(domains) {
  ROAMING["recipientGuard.whitelist.v1"] = { at: Date.now(), emails: [], domains: domains || [] };
}
function setBypass(ageMs) {
  ROAMING["recipientGuard.bypassOnce.v1"] = { at: Date.now() - (ageMs || 0) };
}

var failures = 0;
function run(label, recips, expectAllow, expectContains, expectNotContains, knownList) {
  RECIPS = { to: recips.to || [], cc: recips.cc || [], bcc: recips.bcc || [] };
  setKnown(knownList || []);
  setWhitelist([]); // reset so state can't leak between scenarios
  ROAMING["recipientGuard.bypassOnce.v1"] = undefined;
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
  .then(function () { return run("wrong-person name+prefix clash -> block", { to: [r("Fynn Hodder", "fynn@iteam.je"), r("Fynn Hodder", "fynn@gmail.com")] }, false, "Shares a display name"); })
  .then(function () { return run("external only -> block", { to: [r("Bob", "bob@client.com")] }, false, "Outside your organisation"); })
  .then(function () { return run("personal prefix across domains -> block (prefix clash)", { to: [r("Jon A", "jon.doe@acme.com")], cc: [r("Jon B", "jon.doe@acme-invoices.com")] }, false, "Same username"); })
  // Generic mailbox (sales@) across two real vendors is a legit send, not a
  // wrong-recipient clash: it should NOT produce a "Same username" risk
  // (it's still flagged as external, which is correct).
  .then(function () { return run("generic prefix (sales@) across vendors -> not a prefix clash", { to: [r("Sales A", "sales@acme.com")], cc: [r("Sales B", "sales@partner.com")] }, false, "Outside your organisation", "Same username"); })
  // Regression: a prefix/display-name equal to an Object.prototype key must not
  // throw and fail the guard open (fixed by Object.create(null) grouping maps).
  .then(function () { return run("prototype-key prefix (constructor) -> block", { to: [r("A", "constructor@a.com")], cc: [r("B", "constructor@b.com")] }, false, "Same username"); })
  .then(function () { return run("no recipients -> allow", {}, true); })
  // known-identity (history) checks — the single-wrong-recipient case
  .then(function () {
    return run("single wrong recipient vs known -> block (you usually reach)",
      { to: [r("Fynn Hodder", "fynn.hodder@onecollab.co.uk")] }, false, "You usually use", null,
      [knownRec("Fynn Hodder", "fynn.hodder@gmail.com")]);
  })
  .then(function () {
    // Product decision: flag even when the recipient's own address is known, if
    // the prefix/name resolves to ANOTHER known address.
    // Wording matters here: the recipient IS a known contact, so claiming "you
    // don't usually email this address" would be false. It must say "You ALSO
    // use ..." — the name simply resolves to more than one address they use.
    return run("known recipient with another known same-prefix -> 'You also use' (not 'usually')",
      { to: [r("Fynn Hodder", "fynn.hodder@gmail.com")] }, false, "You also use", "You usually use",
      [knownRec("Fynn Hodder", "fynn.hodder@gmail.com"), knownRec("Fynn Hodder", "fynn.hodder@iteam.je")]);
  })
  .then(function () {
    // But a lone known contact with no same-prefix alternative is NOT flagged as
    // a wrong recipient (only external, since gmail != internal).
    return run("lone known recipient, no alternative -> external only",
      { to: [r("Fynn Hodder", "fynn.hodder@gmail.com")] }, false, "Outside your organisation", "You usually use",
      [knownRec("Fynn Hodder", "fynn.hodder@gmail.com")]);
  })
  .then(function () {
    return run("no known list -> known checks silent (external only)",
      { to: [r("Someone", "someone@onecollab.co.uk")] }, false, "Outside your organisation", "You usually use", []);
  })
  // --- whitelist: a whitelisted address stops producing any risk ---
  .then(function () {
    RECIPS = { to: [r("Bob", "bob@client.com")], cc: [], bcc: [] };
    setKnown([]);
    setWhitelist(["bob@client.com"]);
    ROAMING["recipientGuard.bypassOnce.v1"] = undefined;
    return new Promise(function (resolve) {
      handler({ completed: function (result) { resolve(result); } });
    }).then(function (result) {
      var ok = result.allowEvent === true;
      if (!ok) failures++;
      console.log((ok ? "PASS " : "FAIL ") + "whitelisted external recipient -> allow (no risk)");
    });
  })
  // --- domain whitelist: a whitelisted DOMAIN clears every recipient on it ---
  .then(function () {
    RECIPS = { to: [r("Bob", "bob@client.com")], cc: [r("Sue", "sue@client.com")], bcc: [] };
    setKnown([]);
    setWhitelistDomains(["client.com"]);
    ROAMING["recipientGuard.bypassOnce.v1"] = undefined;
    return new Promise(function (resolve) {
      handler({ completed: function (result) { resolve(result); } });
    }).then(function (result) {
      var ok = result.allowEvent === true; // whole domain trusted -> both external recips allowed
      if (!ok) failures++;
      console.log((ok ? "PASS " : "FAIL ") + "domain-whitelisted -> allow (all recipients on that domain)");
    });
  })
  .then(function () {
    // A domain whitelist for one domain must NOT trust a recipient on another.
    RECIPS = { to: [r("Bob", "bob@client.com")], cc: [], bcc: [] };
    setKnown([]);
    setWhitelistDomains(["partner.com"]);
    ROAMING["recipientGuard.bypassOnce.v1"] = undefined;
    return new Promise(function (resolve) {
      handler({ completed: function (result) { resolve(result); } });
    }).then(function (result) {
      var ok = result.allowEvent === false && (result.errorMessage || "").indexOf("Outside your organisation") !== -1;
      if (!ok) failures++;
      console.log((ok ? "PASS " : "FAIL ") + "domain whitelist is domain-specific (other domains still flag)");
    });
  })
  .then(function () {
    // Whitelisting the wrong address still flags a DIFFERENT wrong recipient.
    RECIPS = { to: [r("Fynn Hodder", "fynn.hodder@onecollab.co.uk")], cc: [], bcc: [] };
    setKnown([knownRec("Fynn Hodder", "fynn.hodder@gmail.com")]);
    setWhitelist(["someone.else@nowhere.com"]);
    ROAMING["recipientGuard.bypassOnce.v1"] = undefined;
    return new Promise(function (resolve) {
      handler({ completed: function (result) { resolve(result); } });
    }).then(function (result) {
      var ok = result.allowEvent === false && (result.errorMessage || "").indexOf("You usually use") !== -1;
      if (!ok) failures++;
      console.log((ok ? "PASS " : "FAIL ") + "unrelated whitelist entry -> still blocks the real wrong recipient");
    });
  })
  // --- one-shot send bypass: a fresh flag releases exactly one send ---
  .then(function () {
    RECIPS = { to: [r("Bob", "bob@client.com")], cc: [], bcc: [] };
    setKnown([]);
    setWhitelist([]);
    setBypass(0); // fresh
    return new Promise(function (resolve) {
      handler({ completed: function (result) { resolve(result); } });
    }).then(function (result) {
      var ok = result.allowEvent === true; // bypass releases this send
      // one-shot: a second send with no fresh flag must block again
      return new Promise(function (resolve) {
        handler({ completed: function (r2) { resolve(r2); } });
      }).then(function (r2) {
        if (r2.allowEvent !== false) ok = false;
        if (!ok) failures++;
        console.log((ok ? "PASS " : "FAIL ") + "fresh bypass releases one send, then re-blocks (one-shot)");
      });
    });
  })
  .then(function () {
    // A stale bypass flag (older than the TTL) must NOT release the send.
    RECIPS = { to: [r("Bob", "bob@client.com")], cc: [], bcc: [] };
    setKnown([]);
    setWhitelist([]);
    setBypass(5 * 60 * 1000); // 5 min old -> stale
    return new Promise(function (resolve) {
      handler({ completed: function (result) { resolve(result); } });
    }).then(function (result) {
      var ok = result.allowEvent === false;
      if (!ok) failures++;
      console.log((ok ? "PASS " : "FAIL ") + "stale bypass flag -> still blocks");
    });
  })
  .then(function () {
    // Smart Alert -> task pane handoff: a block must carry commandId + parseable
    // contextData so the pane can render the full review list.
    RECIPS = { to: [r("Bob", "bob@client.com")], cc: [], bcc: [] };
    setKnown([]);
    setWhitelist([]);
    ROAMING["recipientGuard.bypassOnce.v1"] = undefined;
    return new Promise(function (resolve) {
      handler({ completed: function (result) { resolve(result); } });
    }).then(function (result) {
      var ok = result.allowEvent === false && result.commandId === "RecipientGuard.OpenPane";
      var parsed = null;
      try { parsed = JSON.parse(result.contextData); } catch (e) { /* stays null */ }
      if (!parsed || !parsed.risks || parsed.risks.length === 0) ok = false;
      if (!ok) failures++;
      console.log((ok ? "PASS " : "FAIL ") + "block carries commandId + contextData for the review pane");
    });
  })
  // --- public-domain guard: you must NOT be able to whitelist a whole consumer
  //     domain (that would blind the tool to personal-address mistakes). ---
  .then(function () {
    // Load the task-pane core (exposes RecipientGuardPoc on globalThis in node).
    require(path.join(repoRoot, "src", "recipientGuardCore.js"));
    var RG = globalThis.RecipientGuardPoc;
    setWhitelist([]); // clear
    var ok = RG.isPublicDomain("gmail.com") === true &&
             RG.isPublicDomain("outlook.com") === true &&
             RG.isPublicDomain("acme.com") === false;
    if (!ok) failures++;
    console.log((ok ? "PASS " : "FAIL ") + "isPublicDomain flags consumer domains, not real ones");

    return RG.addDomainToWhitelist("gmail.com").then(function () {
      var blocked = RG.readWhitelist().domains.indexOf("gmail.com") === -1;
      if (!blocked) failures++;
      console.log((blocked ? "PASS " : "FAIL ") + "addDomainToWhitelist REFUSES a public domain (gmail.com)");
      return RG.addDomainToWhitelist("acme.com");
    }).then(function () {
      var added = RG.readWhitelist().domains.indexOf("acme.com") !== -1;
      if (!added) failures++;
      console.log((added ? "PASS " : "FAIL ") + "addDomainToWhitelist allows a real domain (acme.com)");
      return RG.removeFromWhitelist("acme.com");
    }).then(function () {
      var removed = RG.readWhitelist().domains.indexOf("acme.com") === -1;
      if (!removed) failures++;
      console.log((removed ? "PASS " : "FAIL ") + "removeFromWhitelist removes an entry");
      setWhitelist([]);
    });
  })
  .then(function () {
    console.log("");
    if (failures === 0) { console.log("All analysis tests passed."); }
    else { console.log(failures + " test(s) FAILED."); process.exit(1); }
  });
