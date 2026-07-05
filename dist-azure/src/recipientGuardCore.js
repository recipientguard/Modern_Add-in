(function (global) {
  "use strict";

  const RECIPIENT_READ_TIMEOUT_MS = 1200;

  function getMailboxEmail() {
    return (Office.context.mailbox.userProfile.emailAddress || "").trim().toLowerCase();
  }

  function getDomain(email) {
    const value = (email || "").trim().toLowerCase();
    const at = value.lastIndexOf("@");
    return at > -1 ? value.slice(at + 1) : "";
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
            email: (recipient.emailAddress || "").trim().toLowerCase()
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

    return [].concat(groups[0], groups[1], groups[2]).filter((recipient) => recipient.email);
  }

  async function analyzeCurrentMessage() {
    const mailboxEmail = getMailboxEmail();
    const internalDomain = getInternalDomain();
    const recipients = await getAllRecipients();
    const flagged = recipients
      .map((recipient) => {
        const domain = getDomain(recipient.email);
        return {
          type: recipient.type,
          name: recipient.name,
          email: recipient.email,
          domain,
          isExternal: domain !== internalDomain
        };
      })
      .filter((recipient) => recipient.isExternal);

    return {
      mailboxEmail,
      internalDomain,
      recipients,
      flagged,
      hasWarnings: flagged.length > 0
    };
  }

  function buildSmartAlertMessage(analysis) {
    if (!analysis || !analysis.hasWarnings) {
      return "";
    }

    const count = analysis.flagged.length;
    const lines = [
      count === 1
        ? "Recipient Guard found 1 external recipient."
        : "Recipient Guard found " + count + " external recipients.",
      "",
      "Internal domain: " + analysis.internalDomain,
      "",
      "External recipients:"
    ];

    analysis.flagged.slice(0, 8).forEach((recipient) => {
      lines.push("- " + recipient.email);
    });

    if (analysis.flagged.length > 8) {
      lines.push("- +" + (analysis.flagged.length - 8) + " more");
    }

    lines.push("", "Review these recipients before sending.");
    return lines.join("\n");
  }

  global.RecipientGuardPoc = {
    analyzeCurrentMessage,
    buildSmartAlertMessage,
    getInternalDomain,
    getMailboxEmail
  };
})(window);
