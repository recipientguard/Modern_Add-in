// Recipient Guard — shared analysis engine (SINGLE SOURCE OF TRUTH).
//
// This file is the one place the recipient-analysis logic lives. The files the
// manifests actually load (src/sendTestRuntime.js, src/recipientGuardCore.js)
// are GENERATED from it by `npm run build` (scripts/build.js). Edit here, then
// rebuild — never edit the generated files.
//
// Style: ES5/ES2016-safe on purpose (var, function, no ternary, no async/await)
// because classic Outlook on Windows runs event-based add-in JS in a runtime
// that Microsoft documents as failing on newer syntax.

var RECIPIENT_READ_TIMEOUT_MS = 1200;
var MAX_EMAILS_PER_RISK = 6;
var MAX_CONTEXT_RISKS = 20; // cap findings serialized into the Smart Alert contextData
var MAX_RECIPIENTS_IN_ALERT = 8; // cap the recipient list in the system alert text

// The task-pane button the Smart Alert hands off to (must match the <Control>
// id in the manifest). Passing this as commandId in event.completed lets the
// block dialog offer "open the pane" for the full review list.
var PANE_COMMAND_ID = "RecipientGuard.OpenPane";

// --- normalisation helpers (ported from Classic RecipientAnalyzer) ---

function normalizeEmail(value) {
  return (value || "").trim().toLowerCase();
}

function getDomain(email) {
  var value = normalizeEmail(email);
  var at = value.lastIndexOf("@");
  if (at > -1 && at < value.length - 1) {
    return value.slice(at + 1);
  }
  return "";
}

function getLocalPart(email) {
  var value = normalizeEmail(email);
  var at = value.lastIndexOf("@");
  if (at > 0) {
    return value.slice(0, at);
  }
  return "";
}

// Classic Normalize(): lowercase, keep only letters/digits, so punctuation and
// spacing differences still match ("Fynn Hodder" -> "fynnhodder").
function normalizeName(value) {
  return (value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function unique(values) {
  return values.filter(function (value, index) {
    return values.indexOf(value) === index;
  });
}

// --- mailbox / recipient reading ---

function getMailboxEmail() {
  return (Office.context.mailbox.userProfile.emailAddress || "").trim().toLowerCase();
}

function getInternalDomain() {
  return getDomain(getMailboxEmail());
}

function isExternalRecipient(recipient, internalDomain) {
  return Boolean(internalDomain) && recipient.domain !== internalDomain;
}

function getRecipientsAsync(field) {
  return new Promise(function (resolve) {
    var completed = false;
    function finish(value) {
      if (completed) return;
      completed = true;
      resolve(value || []);
    }

    var timer = setTimeout(function () { finish([]); }, RECIPIENT_READ_TIMEOUT_MS);

    try {
      var collection = Office.context.mailbox.item[field];
      if (!collection || typeof collection.getAsync !== "function") {
        clearTimeout(timer);
        finish([]);
        return;
      }

      collection.getAsync(function (result) {
        clearTimeout(timer);
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          finish([]);
          return;
        }
        var recipients = (result.value || []).map(function (recipient) {
          var email = normalizeEmail(recipient.emailAddress);
          return {
            type: field.toUpperCase(),
            name: recipient.displayName || "",
            email: email,
            domain: getDomain(email),
            localPart: getLocalPart(email),
            normalizedName: normalizeName(recipient.displayName)
          };
        });
        finish(recipients);
      });
    } catch (error) {
      clearTimeout(timer);
      finish([]);
    }
  });
}

function getAllRecipients() {
  return Promise.all([
    getRecipientsAsync("to"),
    getRecipientsAsync("cc"),
    getRecipientsAsync("bcc")
  ]).then(function (groups) {
    return [].concat(groups[0], groups[1], groups[2]).filter(function (r) { return r.email; });
  });
}

// --- risk detection ---

function detectExternal(recipients, internalDomain) {
  var risks = [];
  recipients.forEach(function (recipient) {
    if (isExternalRecipient(recipient, internalDomain)) {
      risks.push({ ruleId: "external_domain", severity: "medium", emails: [recipient.email] });
    }
  });
  return risks;
}

// Generic/role mailboxes (sales@, info@, ...) legitimately recur across many
// unrelated domains, so a shared generic prefix is NOT a wrong-recipient
// signal. A personal prefix (fynn, john.smith) shared across domains still is.
var GENERIC_LOCALPARTS = Object.create(null);
[
  "info", "sales", "support", "admin", "administrator", "contact", "hello",
  "team", "accounts", "accounting", "billing", "help", "helpdesk", "office",
  "mail", "enquiries", "inquiries", "marketing", "hr", "jobs", "careers",
  "noreply", "donotreply", "service", "orders", "bookings", "reception",
  "finance", "legal", "press", "media", "newsletter", "webmaster",
  "postmaster", "abuse", "security", "privacy", "feedback", "general"
].forEach(function (name) { GENERIC_LOCALPARTS[name] = true; });

function isGenericLocalPart(localPart) {
  return Boolean(GENERIC_LOCALPARTS[(localPart || "").replace(/[^a-z0-9]/g, "")]);
}

// same_display_name: two recipients share a display name but different addresses.
function detectSameDisplayName(recipients) {
  // Object.create(null): a plain {} would let a display name/prefix of
  // "constructor" or "__proto__" collide with inherited prototype members,
  // throwing and silently failing the check open.
  var byName = Object.create(null);
  recipients.forEach(function (recipient) {
    if (!recipient.normalizedName) return;
    if (!byName[recipient.normalizedName]) byName[recipient.normalizedName] = [];
    byName[recipient.normalizedName].push(recipient);
  });

  var risks = [];
  Object.keys(byName).forEach(function (key) {
    var group = byName[key];
    var emails = unique(group.map(function (r) { return r.email; }));
    if (emails.length > 1) {
      risks.push({ ruleId: "same_display_name", severity: "high", displayName: group[0].name, emails: emails });
    }
  });
  return risks;
}

// same_localpart_different_domain: same prefix before "@", different domains.
function detectSameLocalPart(recipients) {
  var byLocal = Object.create(null); // see note in detectSameDisplayName
  recipients.forEach(function (recipient) {
    if (!recipient.localPart || !recipient.domain) return;
    if (!byLocal[recipient.localPart]) byLocal[recipient.localPart] = [];
    byLocal[recipient.localPart].push(recipient);
  });

  var risks = [];
  Object.keys(byLocal).forEach(function (key) {
    var group = byLocal[key];
    var domains = unique(group.map(function (r) { return r.domain; }));
    if (domains.length > 1 && !isGenericLocalPart(key)) {
      risks.push({
        ruleId: "same_localpart_different_domain",
        severity: "high",
        localPart: key,
        emails: unique(group.map(function (r) { return r.email; }))
      });
    }
  });
  return risks;
}

// --- known-identity store (history-aware detection, from Graph /me/people) ---
//
// The task pane fetches the user's frequently-contacted people and caches a
// compact list in RoamingSettings (per-mailbox, readable from the send-event
// runtime). The send handler and the pane both read it and compare recipients
// against it — this is what catches a SINGLE wrong recipient (no second
// recipient to compare against within the email).

var KNOWN_IDENTITIES_KEY = "recipientGuard.knownIdentities.v1";
var WHITELIST_KEY = "recipientGuard.whitelist.v1";
var BYPASS_KEY = "recipientGuard.bypassOnce.v1";
var BYPASS_TTL_MS = 120000; // a pane-initiated "send now" must be consumed within 2 min

// Compact record: n=normalizedName, e=email, l=localPart, d=domain, name=display.
function toKnownRecord(person) {
  var email = normalizeEmail(person.email || person.emailAddress);
  return {
    name: person.displayName || person.name || "",
    e: email,
    n: normalizeName(person.displayName || person.name),
    l: getLocalPart(email),
    d: getDomain(email)
  };
}

function readKnownIdentities() {
  try {
    var rs = Office.context.roamingSettings;
    if (!rs || typeof rs.get !== "function") return [];
    var stored = rs.get(KNOWN_IDENTITIES_KEY);
    return (stored && stored.people) || [];
  } catch (e) {
    return [];
  }
}

function writeKnownIdentities(records) {
  return new Promise(function (resolve) {
    try {
      var rs = Office.context.roamingSettings;
      rs.set(KNOWN_IDENTITIES_KEY, { at: Date.now(), people: records });
      rs.saveAsync(function () { resolve(true); });
    } catch (e) {
      resolve(false);
    }
  });
}

// --- whitelist (per-address AND per-domain "don't warn about this") ---
//
// Stored in roamingSettings so BOTH the task pane and the send-event runtime see
// it. Shape: { at, emails: [...], domains: [...] }. A recipient is whitelisted if
// its address is in `emails` OR its domain is in `domains`; whitelisted recipients
// are dropped from the analysis input, so they stop producing any risk and stop
// contributing to another recipient's group comparison. A domain entry is the
// clean way to trust a whole partner org — or a second internal domain, which
// also sidesteps the single-internal-domain limitation.
//
// Backward compatible: older data was { at, emails } with no domains.

function normalizeDomain(value) {
  var v = (value || "").trim().toLowerCase();
  var at = v.lastIndexOf("@");
  if (at !== -1) v = v.slice(at + 1); // tolerate "user@acme.com" or "@acme.com"
  return v;
}

// Public / consumer mail domains must NEVER be domain-whitelisted — trusting the
// whole of gmail.com would blind the add-in to exactly the personal-address
// mistakes it exists to catch. The UI hides the "trust this domain" action for
// these, and addDomainToWhitelist refuses them as a backstop; per-ADDRESS
// whitelisting still works.
var PUBLIC_EMAIL_DOMAINS = Object.create(null);
[
  "gmail.com", "googlemail.com",
  "outlook.com", "outlook.co.uk", "hotmail.com", "hotmail.co.uk", "live.com",
  "live.co.uk", "msn.com",
  "yahoo.com", "yahoo.co.uk", "ymail.com", "rocketmail.com",
  "icloud.com", "me.com", "mac.com",
  "aol.com", "gmx.com", "gmx.net", "mail.com",
  "proton.me", "protonmail.com", "pm.me",
  "zoho.com", "yandex.com", "yandex.ru", "qq.com", "163.com", "126.com"
].forEach(function (d) { PUBLIC_EMAIL_DOMAINS[d] = true; });

function isPublicDomain(domain) {
  return Boolean(PUBLIC_EMAIL_DOMAINS[normalizeDomain(domain)]);
}

function readWhitelist() {
  try {
    var rs = Office.context.roamingSettings;
    if (!rs || typeof rs.get !== "function") return { emails: [], domains: [] };
    var stored = rs.get(WHITELIST_KEY) || {};
    return { emails: stored.emails || [], domains: stored.domains || [] };
  } catch (e) {
    return { emails: [], domains: [] };
  }
}

function writeWhitelist(wl) {
  return new Promise(function (resolve) {
    var value = { at: Date.now(), emails: (wl && wl.emails) || [], domains: (wl && wl.domains) || [] };
    try {
      var rs = Office.context.roamingSettings;
      rs.set(WHITELIST_KEY, value);
      rs.saveAsync(function () { resolve(value); });
    } catch (e) {
      resolve(readWhitelist());
    }
  });
}

function isWhitelisted(email, whitelist) {
  var wl = whitelist || { emails: [], domains: [] };
  var e = normalizeEmail(email);
  if ((wl.emails || []).indexOf(e) !== -1) return true;
  var d = getDomain(e);
  return Boolean(d) && (wl.domains || []).indexOf(d) !== -1;
}

function addToWhitelist(email) {
  var wl = readWhitelist();
  var e = normalizeEmail(email);
  if (e && wl.emails.indexOf(e) === -1) wl.emails = wl.emails.concat([e]);
  return writeWhitelist(wl);
}

function addDomainToWhitelist(domain) {
  var wl = readWhitelist();
  var d = normalizeDomain(domain);
  // Never allow a public/consumer domain, even if the UI somehow asked.
  if (d && !isPublicDomain(d) && wl.domains.indexOf(d) === -1) {
    wl.domains = wl.domains.concat([d]);
  }
  return writeWhitelist(wl);
}

// Remove an entry from either list (used by the pane's manage view). Matches a
// bare address or domain string.
function removeFromWhitelist(value) {
  var wl = readWhitelist();
  var v = (value || "").trim().toLowerCase();
  wl.emails = wl.emails.filter(function (x) { return x !== v; });
  wl.domains = wl.domains.filter(function (x) { return x !== v; });
  return writeWhitelist(wl);
}

function excludeWhitelisted(recipients, whitelist) {
  var wl = whitelist || { emails: [], domains: [] };
  if ((wl.emails || []).length === 0 && (wl.domains || []).length === 0) return recipients;
  return recipients.filter(function (r) { return !isWhitelisted(r.email, wl); });
}

// --- one-shot send bypass (pane "send now" past a block) ---
//
// sendAsync from the task pane re-triggers OnMessageSend (separate runtimes), so
// a pane-initiated send would re-block. The pane sets a fresh, short-lived flag;
// the send handler consumes it once and lets that single send through.

function setSendBypass() {
  return new Promise(function (resolve) {
    try {
      var rs = Office.context.roamingSettings;
      rs.set(BYPASS_KEY, { at: Date.now() });
      rs.saveAsync(function () { resolve(true); });
    } catch (e) {
      resolve(false);
    }
  });
}

function clearSendBypass() {
  return new Promise(function (resolve) {
    try {
      var rs = Office.context.roamingSettings;
      rs.set(BYPASS_KEY, undefined);
      rs.saveAsync(function () { resolve(true); });
    } catch (e) {
      resolve(false);
    }
  });
}

// Returns true if a fresh pane-initiated bypass is present, and clears it so it
// can only ever release one send (one-shot). Stale flags are ignored.
function consumeSendBypass() {
  try {
    var rs = Office.context.roamingSettings;
    if (!rs || typeof rs.get !== "function") return false;
    var flag = rs.get(BYPASS_KEY);
    var fresh = Boolean(flag && flag.at && (Date.now() - flag.at) < BYPASS_TTL_MS);
    if (flag) {
      try { rs.set(BYPASS_KEY, undefined); rs.saveAsync(function () {}); } catch (e2) { /* best effort */ }
    }
    return fresh;
  } catch (e) {
    return false;
  }
}

// known_display_name / known_localpart: a recipient's name or prefix matches
// someone you usually email at a DIFFERENT address. By design we flag this even
// when the recipient's own address is also a known contact — if a name/prefix
// resolves to more than one address you use, we surface it and let the user
// decide (better a dismissed warning than a wrong pick). A known identity with
// the exact same email as the recipient is still excluded (that's not an
// alternative).
function detectKnownAlternatives(recipients, knownIdentities) {
  var risks = [];
  if (!knownIdentities || knownIdentities.length === 0) return risks;

  recipients.forEach(function (recipient) {
    // Collect ALL known alternatives for this recipient — matched by display
    // name and/or email name — de-duplicated, tracking WHY each one matched so
    // the message can annotate each alternative individually.
    var altReasons = Object.create(null); // email -> { name: bool, prefix: bool }

    knownIdentities.forEach(function (k) {
      if (k.e === recipient.email) return; // same address: not an alternative
      var byName = Boolean(recipient.normalizedName && k.n && k.n === recipient.normalizedName);
      var byPrefix = Boolean(recipient.localPart && recipient.domain && k.l && k.l === recipient.localPart && k.d && k.d !== recipient.domain);
      if (!byName && !byPrefix) return;
      if (!altReasons[k.e]) altReasons[k.e] = { name: false, prefix: false };
      if (byName) altReasons[k.e].name = true;
      if (byPrefix) altReasons[k.e].prefix = true;
    });

    var alternatives = Object.keys(altReasons).map(function (email) {
      return { email: email, byName: altReasons[email].name, byPrefix: altReasons[email].prefix };
    });
    if (alternatives.length > 0) {
      // Is the recipient's OWN address one we already know? If so the flag isn't
      // "you don't email this person" — it's "this name resolves to several
      // addresses you use, check you picked the right one". The wording differs,
      // so record which case this is.
      var recipientIsKnown = knownIdentities.some(function (k) { return k.e === recipient.email; });
      risks.push({
        ruleId: "known_alternative",
        severity: "high",
        emails: [recipient.email],
        recipientIsKnown: recipientIsKnown,
        alternatives: alternatives
      });
    }
  });
  return risks;
}

function isStrong(risk) {
  return risk.ruleId === "same_display_name" ||
    risk.ruleId === "same_localpart_different_domain" ||
    risk.ruleId === "known_alternative";
}

// Condense (ported from RiskSignalOrdering.Condense): if an address is already
// implicated by a strong signal, don't also report it as merely external.
function condense(risks) {
  var strongEmails = Object.create(null);
  risks.filter(isStrong).forEach(function (risk) {
    risk.emails.forEach(function (email) { strongEmails[email] = true; });
  });
  return risks.filter(function (risk) {
    if (risk.ruleId !== "external_domain") return true;
    return !risk.emails.every(function (email) { return strongEmails[email]; });
  });
}

function computeRisks(recipients, internalDomain, knownIdentities, whitelist) {
  var input = excludeWhitelisted(recipients, whitelist);
  var risks = detectSameDisplayName(input)
    .concat(detectSameLocalPart(input))
    .concat(detectKnownAlternatives(input, knownIdentities || []))
    .concat(detectExternal(input, internalDomain));
  return condense(risks);
}

// --- Smart Alert / task-pane message ---

// Shared reason wording for a flagged recipient — used by the send-time alert,
// the review dialog, and the pane so every surface explains a flag identically.
//
// For known_alternative we name ONE alternative (the highest-ranked match, since
// /me/people is relevance-ordered) rather than listing them all — the full list
// was too noisy, but a bare reason carried no meaning. Wording depends on whether
// the recipient's own address is already known:
//   not known -> "You usually use x@y"      (they've probably picked the wrong one)
//   known     -> "You also use x@y"         (the name has several addresses; which is right?)
// Saying "you don't usually email this address" about an address that IS in their
// contacts is simply false, and reads as a broken warning.
function noteForRisk(risk) {
  if (!risk) return "";
  if (risk.ruleId === "known_alternative") {
    var alts = risk.alternatives || [];
    if (alts.length === 0) return "You don't usually email this address";
    var note = (risk.recipientIsKnown ? "You also use " : "You usually use ") + alts[0].email;
    if (alts.length > 1) note += " (+" + (alts.length - 1) + " more)";
    return note;
  }
  if (risk.ruleId === "same_display_name") return "Shares a display name with another recipient";
  if (risk.ruleId === "same_localpart_different_domain") return "Same username as another recipient";
  if (risk.ruleId === "external_domain") return "Outside your organisation";
  return "";
}

function formatRecipientType(type) {
  var s = (type || "").toLowerCase();
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : "";
}

function buildAlertMessage(risks, recipients) {
  if (!risks || risks.length === 0) return "";

  // One reason per flagged address (first/strongest wins; condense() prevents a
  // real overlap between external and the stronger signals).
  var noteByEmail = Object.create(null);
  var order = [];
  risks.forEach(function (risk) {
    var note = noteForRisk(risk);
    if (!note) return;
    (risk.emails || []).forEach(function (email) {
      if (!noteByEmail[email]) { noteByEmail[email] = note; order.push(email); }
    });
  });

  var lines = ["Recipient Guard paused this send. Please check the recipients are correct.", "", "This message will be sent to:"];

  // List EVERY recipient (matching the review dialog), flagged ones annotated.
  // Fall back to just the flagged addresses if the recipient list is unavailable.
  var list = (recipients && recipients.length) ? recipients : null;
  if (list) {
    list.slice(0, MAX_RECIPIENTS_IN_ALERT).forEach(function (r) {
      var t = formatRecipientType(r.type);
      lines.push("  " + r.email + (t ? "  (" + t + ")" : ""));
      if (noteByEmail[r.email]) lines.push("    " + noteByEmail[r.email]);
    });
    if (list.length > MAX_RECIPIENTS_IN_ALERT) {
      lines.push("  +" + (list.length - MAX_RECIPIENTS_IN_ALERT) + " more");
    }
  } else {
    order.slice(0, MAX_RECIPIENTS_IN_ALERT).forEach(function (email) {
      lines.push("  " + email);
      lines.push("    " + noteByEmail[email]);
    });
  }

  lines.push("");
  lines.push("Choose Take action to review, or Send anyway to send as is.");
  return lines.join("\n");
}

// Serialize the findings for the Smart Alert -> task pane handoff. Passed as
// contextData in event.completed and re-hydrated by the pane via
// getInitializationContextAsync to render the full review list. Kept compact and
// capped so it stays well under the platform's contextData size limit.
function buildContextData(risks) {
  var trimmed = (risks || []).slice(0, MAX_CONTEXT_RISKS).map(function (risk) {
    var out = { ruleId: risk.ruleId, emails: (risk.emails || []).slice(0, MAX_EMAILS_PER_RISK) };
    if (risk.displayName) out.displayName = risk.displayName;
    if (risk.localPart) out.localPart = risk.localPart;
    if (risk.alternatives) {
      out.alternatives = risk.alternatives.slice(0, MAX_EMAILS_PER_RISK).map(function (a) {
        return { email: a.email, byName: a.byName, byPrefix: a.byPrefix };
      });
    }
    return out;
  });
  return JSON.stringify({ v: 1, risks: trimmed });
}
