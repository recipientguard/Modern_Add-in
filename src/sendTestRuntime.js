(function () {
  "use strict";

  // Self-contained OnMessageSend runtime. Reads the recipients directly at send
  // time and runs the analysis inline — no dependency on the task pane, no
  // localStorage handoff (which isn't shared across separate runtimes), and no
  // "click Check recipients first" step. The engine below is duplicated from
  // recipientGuardCore.js on purpose; a later build step will bundle the shared
  // core to remove the copy.
  //
  // Kept ES2016-safe (no async/await, no ternary operator) so it also loads in
  // older classic Outlook on Windows, which Microsoft documents as failing on
  // those constructs in event runtimes.

  var RECIPIENT_READ_TIMEOUT_MS = 1200;
  var SEND_SAFETY_TIMEOUT_MS = 3000; // fail-open well under Outlook's 5s limit
  var MAX_EMAILS_PER_RISK = 6;

  // --- helpers ---

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

  function normalizeName(value) {
    return (value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function unique(values) {
    return values.filter(function (value, index) {
      return values.indexOf(value) === index;
    });
  }

  // --- mailbox / recipient reading (inline at send time) ---

  function getInternalDomain() {
    return getDomain((Office.context.mailbox.userProfile.emailAddress || "").trim());
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
      var isExternal = Boolean(internalDomain) && recipient.domain !== internalDomain;
      if (isExternal) {
        risks.push({ ruleId: "external_domain", severity: "medium", emails: [recipient.email] });
      }
    });
    return risks;
  }

  function detectSameDisplayName(recipients) {
    var byName = {};
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

  function detectSameLocalPart(recipients) {
    var byLocal = {};
    recipients.forEach(function (recipient) {
      if (!recipient.localPart || !recipient.domain) return;
      if (!byLocal[recipient.localPart]) byLocal[recipient.localPart] = [];
      byLocal[recipient.localPart].push(recipient);
    });

    var risks = [];
    Object.keys(byLocal).forEach(function (key) {
      var group = byLocal[key];
      var domains = unique(group.map(function (r) { return r.domain; }));
      if (domains.length > 1) {
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

  function condense(risks) {
    var strongEmails = {};
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

  // --- message ---

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

  // --- send event handler ---

  function onMessageSendDiagnostic(event) {
    var settled = false;
    function finish(result) {
      if (settled) return;
      settled = true;
      try { event.completed(result); } catch (ignored) { /* nothing more we can do */ }
    }

    // Never let the send hang: if analysis stalls, fail open (allow the send).
    var safety = setTimeout(function () { finish({ allowEvent: true }); }, SEND_SAFETY_TIMEOUT_MS);

    try {
      var internalDomain = getInternalDomain();
      getAllRecipients().then(function (recipients) {
        clearTimeout(safety);
        var risks = computeRisks(recipients, internalDomain);
        if (risks.length === 0) {
          finish({ allowEvent: true });
        } else {
          finish({ allowEvent: false, errorMessage: buildAlertMessage(risks) });
        }
      })["catch"](function () {
        clearTimeout(safety);
        finish({ allowEvent: true });
      });
    } catch (error) {
      clearTimeout(safety);
      finish({ allowEvent: true });
    }
  }

  Office.actions.associate("onMessageSendDiagnostic", onMessageSendDiagnostic);
  if (typeof globalThis !== "undefined") {
    globalThis.onMessageSendDiagnostic = onMessageSendDiagnostic;
  }
})();
