(function () {
  "use strict";

  var ANALYSIS_CACHE_KEY = "recipientGuard.latestComposeAnalysis.v1";
  var ANALYSIS_MAX_AGE_MS = 5 * 60 * 1000;
  var MAX_RECIPIENTS_IN_ALERT = 6;

  function normalizeEmail(value) {
    return (value || "").trim().toLowerCase();
  }

  function getDomain(email) {
    var value = normalizeEmail(email);
    var at = value.lastIndexOf("@");
    return at > -1 ? value.slice(at + 1) : "";
  }

  function readCachedAnalysis() {
    var raw = localStorage.getItem(ANALYSIS_CACHE_KEY);
    if (!raw) return null;

    var analysis = JSON.parse(raw);
    if (!analysis || !analysis.createdAt) return null;
    if ((Date.now() - analysis.createdAt) > ANALYSIS_MAX_AGE_MS) return null;

    return analysis;
  }

  function buildExternalRecipientMessage(internalDomain, externalRecipients) {
    var count = externalRecipients.length;
    var lines = [
      count === 1
        ? "Recipient Guard found 1 external recipient."
        : "Recipient Guard found " + count + " external recipients.",
      "",
      "Internal domain: " + internalDomain,
      "",
      count === 1 ? "External recipient:" : "External recipients:"
    ];

    externalRecipients.slice(0, MAX_RECIPIENTS_IN_ALERT).forEach(function (recipient) {
      lines.push(recipient.email);
    });

    if (count > MAX_RECIPIENTS_IN_ALERT) {
      lines.push("+" + (count - MAX_RECIPIENTS_IN_ALERT) + " more");
    }

    lines.push("", "Review these recipients before sending.");
    return lines.join("\n");
  }

  function completeSend(event, result) {
    event.completed(result);
  }

  function blockForMissingCheck(event) {
    completeSend(event, {
      allowEvent: false,
      errorMessage: "Recipient Guard needs a fresh recipient check. Open Recipient Guard Send Test from Apps, click Check recipients, then send again."
    });
  }

  function onMessageSendDiagnostic(event) {
    try {
      var analysis = readCachedAnalysis();
      if (!analysis || !analysis.internalDomain || !Array.isArray(analysis.recipients)) {
        blockForMissingCheck(event);
        return;
      }

      var externalRecipients = (analysis.flagged || []).filter(function (recipient) {
        return recipient && recipient.email && getDomain(recipient.email) !== analysis.internalDomain;
      });

      if (externalRecipients.length === 0) {
        completeSend(event, { allowEvent: true });
        return;
      }

      completeSend(event, {
        allowEvent: false,
        errorMessage: buildExternalRecipientMessage(analysis.internalDomain, externalRecipients)
      });
    } catch (error) {
      blockForMissingCheck(event);
    }
  }

  globalThis.onMessageSendDiagnostic = onMessageSendDiagnostic;

  if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendDiagnostic", onMessageSendDiagnostic);
  }
})();
