(function (global) {
  "use strict";

  const RECIPIENT_READ_TIMEOUT_MS = 1200;
  const MAX_EMAILS_PER_RISK = 6;

  // --- normalisation helpers (ported from Classic RecipientAnalyzer) ---

  function normalizeEmail(value) {
    return (value || "").trim().toLowerCase();
  }

  function getDomain(email) {
    const value = normalizeEmail(email);
    const at = value.lastIndexOf("@");
    return at > -1 && at < value.length - 1 ? value.slice(at + 1) : "";
  }

  function getLocalPart(email) {
    const value = normalizeEmail(email);
    const at = value.lastIndexOf("@");
    return at > 0 ? value.slice(0, at) : "";
  }

  // Classic Normalize(): lowercase, keep only letters/digits.
  // "Fynn Hodder" -> "fynnhodder", so punctuation/spacing differences still match.
  function normalizeName(value) {
    return (value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function unique(values) {
    return values.filter((value, index) => values.indexOf(value) === index);
  }

  // --- mailbox / recipient reading ---

  function getMailboxEmail() {
    return (Office.context.mailbox.userProfile.emailAddress || "").trim().toLowerCase();
  }

  function getInternalDomain() {
    return getDomain(getMailboxEmail());
  }

  function getRecipientsAsync(field) {
    return new Promise((resolve) => {
      let completed = false;
      const finish = (value) => {
        if (completed) return;
        completed = true;
        resolve(value || []);
      };

      const timer = setTimeout(() => finish([]), RECIPIENT_READ_TIMEOUT_MS);

      try {
        const collection = Office.context.mailbox.item[field];
        if (!collection || typeof collection.getAsync !== "function") {
          clearTimeout(timer);
          finish([]);
          return;
        }

        collection.getAsync((result) => {
          clearTimeout(timer);
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            finish([]);
            return;
          }

          const recipients = (result.value || []).map((recipient) => ({
            type: field.toUpperCase(),
            name: recipient.displayName || "",
            email: normalizeEmail(recipient.emailAddress)
          }));
          finish(recipients);
        });
      } catch (error) {
        clearTimeout(timer);
        finish([]);
      }
    });
  }

  async function getAllRecipients() {
    const groups = await Promise.all([
      getRecipientsAsync("to"),
      getRecipientsAsync("cc"),
      getRecipientsAsync("bcc")
    ]);

    return [].concat(groups[0], groups[1], groups[2])
      .filter((recipient) => recipient.email)
      .map((recipient) => ({
        type: recipient.type,
        name: recipient.name,
        email: recipient.email,
        domain: getDomain(recipient.email),
        localPart: getLocalPart(recipient.email),
        normalizedName: normalizeName(recipient.name)
      }));
  }

  // --- risk detection (history-free: current recipients only) ---

  function detectExternal(recipients, internalDomain) {
    const risks = [];
    recipients.forEach((recipient) => {
      recipient.isExternal = Boolean(internalDomain) && recipient.domain !== internalDomain;
      if (recipient.isExternal) {
        risks.push({
          ruleId: "external_domain",
          severity: "medium",
          emails: [recipient.email]
        });
      }
    });
    return risks;
  }

  // same_display_name: two recipients share a display name but use different addresses.
  function detectSameDisplayName(recipients) {
    const byName = {};
    recipients.forEach((recipient) => {
      if (!recipient.normalizedName) return;
      (byName[recipient.normalizedName] = byName[recipient.normalizedName] || []).push(recipient);
    });

    return Object.keys(byName).reduce((risks, key) => {
      const group = byName[key];
      const emails = unique(group.map((recipient) => recipient.email));
      if (emails.length > 1) {
        risks.push({
          ruleId: "same_display_name",
          severity: "high",
          displayName: group[0].name,
          emails
        });
      }
      return risks;
    }, []);
  }

  // same_localpart_different_domain: same prefix before "@", different domains.
  function detectSameLocalPart(recipients) {
    const byLocal = {};
    recipients.forEach((recipient) => {
      if (!recipient.localPart || !recipient.domain) return;
      (byLocal[recipient.localPart] = byLocal[recipient.localPart] || []).push(recipient);
    });

    return Object.keys(byLocal).reduce((risks, key) => {
      const group = byLocal[key];
      const domains = unique(group.map((recipient) => recipient.domain));
      if (domains.length > 1) {
        risks.push({
          ruleId: "same_localpart_different_domain",
          severity: "high",
          localPart: key,
          emails: unique(group.map((recipient) => recipient.email))
        });
      }
      return risks;
    }, []);
  }

  function isStrong(risk) {
    return risk.ruleId === "same_display_name" || risk.ruleId === "same_localpart_different_domain";
  }

  // Condense (ported from RiskSignalOrdering.Condense): if an address is already
  // implicated by a strong signal, don't also report it as merely external.
  function condense(risks) {
    const strongEmails = {};
    risks.filter(isStrong).forEach((risk) => {
      risk.emails.forEach((email) => { strongEmails[email] = true; });
    });

    return risks.filter((risk) => {
      if (risk.ruleId !== "external_domain") return true;
      return !risk.emails.every((email) => strongEmails[email]);
    });
  }

  function computeRisks(recipients, internalDomain) {
    const risks = []
      .concat(detectSameDisplayName(recipients))
      .concat(detectSameLocalPart(recipients))
      .concat(detectExternal(recipients, internalDomain));
    return condense(risks);
  }

  async function analyzeCurrentMessage() {
    const mailboxEmail = getMailboxEmail();
    const internalDomain = getInternalDomain();
    const recipients = await getAllRecipients();
    const risks = computeRisks(recipients, internalDomain);

    return {
      mailboxEmail,
      internalDomain,
      recipients,
      risks,
      // Kept for the task-pane's per-recipient external badges.
      flagged: recipients.filter((recipient) => recipient.isExternal),
      hasWarnings: risks.length > 0
    };
  }

  // --- Smart Alert message (shared by task pane and send runtime) ---

  function listEmails(lines, emails) {
    emails.slice(0, MAX_EMAILS_PER_RISK).forEach((email) => lines.push("  " + email));
    if (emails.length > MAX_EMAILS_PER_RISK) {
      lines.push("  +" + (emails.length - MAX_EMAILS_PER_RISK) + " more");
    }
  }

  function buildAlertMessage(analysis) {
    if (!analysis || !Array.isArray(analysis.risks) || analysis.risks.length === 0) {
      return "";
    }

    const risks = analysis.risks;
    const hasStrong = risks.some(isStrong);
    const lines = [
      hasStrong
        ? "Recipient Guard found a possible wrong recipient."
        : (risks.length === 1
            ? "Recipient Guard found 1 external recipient."
            : "Recipient Guard found " + risks.length + " external recipients."),
      ""
    ];

    risks.filter((risk) => risk.ruleId === "same_display_name").forEach((risk) => {
      lines.push('Recipients share the name "' + (risk.displayName || "").trim() + '" but use different addresses:');
      listEmails(lines, risk.emails);
      lines.push("");
    });

    risks.filter((risk) => risk.ruleId === "same_localpart_different_domain").forEach((risk) => {
      lines.push('Same address prefix "' + risk.localPart + '" on different domains:');
      listEmails(lines, risk.emails);
      lines.push("");
    });

    const external = risks.filter((risk) => risk.ruleId === "external_domain");
    if (external.length > 0) {
      lines.push(external.length === 1 ? "External recipient:" : "External recipients:");
      listEmails(lines, external.map((risk) => risk.emails[0]));
      lines.push("");
    }

    lines.push("Review these recipients before sending.");
    return lines.join("\n");
  }

  global.RecipientGuardPoc = {
    analyzeCurrentMessage,
    buildAlertMessage,
    computeRisks,
    getInternalDomain,
    getMailboxEmail,
    // exposed for reuse/testing
    normalizeName,
    getLocalPart,
    getDomain
  };
})(window);
