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

// --- whitelist (per-address "don't warn about this again") ---
//
// Stored in roamingSettings so BOTH the task pane and the send-event runtime see
// it. A whitelisted address is dropped from the analysis input, so it stops
// producing any risk (external / known-alternative / group) and also stops
// contributing to another recipient's group comparison.

function readWhitelist() {
  try {
    var rs = Office.context.roamingSettings;
    if (!rs || typeof rs.get !== "function") return [];
    var stored = rs.get(WHITELIST_KEY);
    return (stored && stored.emails) || [];
  } catch (e) {
    return [];
  }
}

function isWhitelisted(email, whitelist) {
  return (whitelist || []).indexOf(normalizeEmail(email)) !== -1;
}

function addToWhitelist(email) {
  return new Promise(function (resolve) {
    try {
      var rs = Office.context.roamingSettings;
      var current = readWhitelist();
      var e = normalizeEmail(email);
      if (current.indexOf(e) === -1) current = current.concat([e]);
      rs.set(WHITELIST_KEY, { at: Date.now(), emails: current });
      rs.saveAsync(function () { resolve(current); });
    } catch (err) {
      resolve(readWhitelist());
    }
  });
}

function excludeWhitelisted(recipients, whitelist) {
  if (!whitelist || whitelist.length === 0) return recipients;
  return recipients.filter(function (r) { return !isWhitelisted(r.email, whitelist); });
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
      risks.push({
        ruleId: "known_alternative",
        severity: "high",
        emails: [recipient.email],
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

// One place for the "why this is an alternative" wording, shared by the send
// dialog message and the task-pane review list. Returns bare text; callers add
// their own punctuation/parentheses.
function describeAlternative(alt) {
  if (alt.byName && alt.byPrefix) return "same display name & username";
  if (alt.byName) return "same display name";
  return "same username";
}

function listEmails(lines, emails) {
  emails.slice(0, MAX_EMAILS_PER_RISK).forEach(function (email) { lines.push("  " + email); });
  if (emails.length > MAX_EMAILS_PER_RISK) {
    lines.push("  +" + (emails.length - MAX_EMAILS_PER_RISK) + " more");
  }
}

function buildAlertMessage(risks) {
  if (!risks || risks.length === 0) return "";

  var hasKnown = risks.some(function (r) { return r.ruleId === "known_alternative"; });
  var hasStrong = risks.some(isStrong);
  var header;
  if (hasKnown) {
    header = "Recipient Guard found a recipient that may have been picked incorrectly from AutoComplete.";
  } else if (hasStrong) {
    header = "Recipient Guard found a possible wrong recipient.";
  } else if (risks.length === 1) {
    header = "Recipient Guard found 1 external recipient.";
  } else {
    header = "Recipient Guard found " + risks.length + " external recipients.";
  }
  var lines = [header, ""];

  risks.filter(function (r) { return r.ruleId === "known_alternative"; }).forEach(function (r) {
    lines.push("Sending to: " + r.emails[0]);
    lines.push("You usually use:");
    r.alternatives.slice(0, MAX_EMAILS_PER_RISK).forEach(function (alt) {
      lines.push("  " + alt.email + "  (" + describeAlternative(alt) + ")");
    });
    if (r.alternatives.length > MAX_EMAILS_PER_RISK) {
      lines.push("  +" + (r.alternatives.length - MAX_EMAILS_PER_RISK) + " more");
    }
    lines.push("");
  });

  risks.filter(function (r) { return r.ruleId === "same_display_name"; }).forEach(function (r) {
    lines.push('Recipients share the display name "' + (r.displayName || "").trim() + '" but use different addresses:');
    listEmails(lines, r.emails);
    lines.push("");
  });

  risks.filter(function (r) { return r.ruleId === "same_localpart_different_domain"; }).forEach(function (r) {
    lines.push('Same username "' + r.localPart + '" on different domains:');
    listEmails(lines, r.emails);
    lines.push("");
  });

  var external = risks.filter(function (r) { return r.ruleId === "external_domain"; });
  if (external.length > 0) {
    if (external.length === 1) {
      lines.push("External recipient:");
    } else {
      lines.push("External recipients:");
    }
    listEmails(lines, external.map(function (r) { return r.emails[0]; }));
    lines.push("");
  }

  lines.push("Review these recipients before sending.");
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
