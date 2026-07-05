(function () {
  "use strict";

  var ANALYSIS_CACHE_KEY = "recipientGuard.latestComposeAnalysis.v1";
  var ANALYSIS_MAX_AGE_MS = 5 * 60 * 1000;

  function readCachedAnalysis() {
    try {
      var raw = localStorage.getItem(ANALYSIS_CACHE_KEY);
      if (!raw) return null;

      var analysis = JSON.parse(raw);
      if (!analysis || !analysis.createdAt) return null;
      if ((Date.now() - analysis.createdAt) > ANALYSIS_MAX_AGE_MS) return null;

      return analysis;
    } catch (error) {
      return null;
    }
  }

  function completeSend(event, result) {
    event.completed(result);
  }

  function blockForMissingCheck(event) {
    completeSend(event, {
      allowEvent: false,
      errorMessage: "Recipient Guard needs a fresh recipient check. Open Recipient Guard from Apps, click Check recipients, then send again."
    });
  }

  function onMessageSendDiagnostic(event) {
    try {
      var analysis = readCachedAnalysis();
      if (!analysis || !Array.isArray(analysis.risks)) {
        blockForMissingCheck(event);
        return;
      }

      if (analysis.risks.length === 0) {
        completeSend(event, { allowEvent: true });
        return;
      }

      var message = window.RecipientGuardPoc && window.RecipientGuardPoc.buildAlertMessage
        ? window.RecipientGuardPoc.buildAlertMessage(analysis)
        : "Recipient Guard found a possible recipient issue. Review the recipients before sending.";

      completeSend(event, { allowEvent: false, errorMessage: message });
    } catch (error) {
      blockForMissingCheck(event);
    }
  }

  globalThis.onMessageSendDiagnostic = onMessageSendDiagnostic;

  if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendDiagnostic", onMessageSendDiagnostic);
  }
})();
