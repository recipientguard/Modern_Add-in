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

function isStrong(risk) {
  return risk.ruleId === "same_display_name" || risk.ruleId === "same_localpart_different_domain";
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

function computeRisks(recipients, internalDomain) {
  var risks = detectSameDisplayName(recipients)
    .concat(detectSameLocalPart(recipients))
    .concat(detectExternal(recipients, internalDomain));
  return condense(risks);
}

// --- Smart Alert / task-pane message ---

function listEmails(lines, emails) {
  emails.slice(0, MAX_EMAILS_PER_RISK).forEach(function (email) { lines.push("  " + email); });
  if (emails.length > MAX_EMAILS_PER_RISK) {
    lines.push("  +" + (emails.length - MAX_EMAILS_PER_RISK) + " more");
  }
}

function buildAlertMessage(risks) {
  if (!risks || risks.length === 0) return "";

  var hasStrong = risks.some(isStrong);
  var header;
  if (hasStrong) {
    header = "Recipient Guard found a possible wrong recipient.";
  } else if (risks.length === 1) {
    header = "Recipient Guard found 1 external recipient.";
  } else {
    header = "Recipient Guard found " + risks.length + " external recipients.";
  }
  var lines = [header, ""];

  risks.filter(function (r) { return r.ruleId === "same_display_name"; }).forEach(function (r) {
    lines.push('Recipients share the name "' + (r.displayName || "").trim() + '" but use different addresses:');
    listEmails(lines, r.emails);
    lines.push("");
  });

  risks.filter(function (r) { return r.ruleId === "same_localpart_different_domain"; }).forEach(function (r) {
    lines.push('Same address prefix "' + r.localPart + '" on different domains:');
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
