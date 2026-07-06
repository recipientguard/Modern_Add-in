(function () {
  "use strict";

  // Self-contained send-event runtime. On new Outlook / OWA the event runs as a
  // JavaScript-only runtime (manifest <Override type="javascript">) and does NOT
  // load recipientGuardCore.js, so this file must not depend on it. The message
  // builder below is duplicated from recipientGuardCore.js on purpose; a later
  // build step will bundle the shared core into this runtime to remove the copy.

  var ANALYSIS_CACHE_KEY = "recipientGuard.latestComposeAnalysis.v1";
  var ANALYSIS_MAX_AGE_MS = 5 * 60 * 1000;
  var MAX_EMAILS_PER_RISK = 6;

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

  function isStrong(risk) {
    return risk.ruleId === "same_display_name" || risk.ruleId === "same_localpart_different_domain";
  }

  function listEmails(lines, emails) {
    emails.slice(0, MAX_EMAILS_PER_RISK).forEach(function (email) { lines.push("  " + email); });
    if (emails.length > MAX_EMAILS_PER_RISK) {
      lines.push("  +" + (emails.length - MAX_EMAILS_PER_RISK) + " more");
    }
  }

  function buildAlertMessage(analysis) {
    if (!analysis || !Array.isArray(analysis.risks) || analysis.risks.length === 0) return "";

    var risks = analysis.risks;
    var hasStrong = risks.some(isStrong);
    var lines = [
      hasStrong
        ? "Recipient Guard found a possible wrong recipient."
        : (risks.length === 1
            ? "Recipient Guard found 1 external recipient."
            : "Recipient Guard found " + risks.length + " external recipients."),
      ""
    ];

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
      lines.push(external.length === 1 ? "External recipient:" : "External recipients:");
      listEmails(lines, external.map(function (r) { return r.emails[0]; }));
      lines.push("");
    }

    lines.push("Review these recipients before sending.");
    return lines.join("\n");
  }

  function onMessageSendDiagnostic(event) {
    try {
      var analysis = readCachedAnalysis();
      if (!analysis || !Array.isArray(analysis.risks)) {
        event.completed({
          allowEvent: false,
          errorMessage: "Recipient Guard needs a fresh recipient check. Open Recipient Guard from Apps, click Check recipients, then send again."
        });
        return;
      }

      if (analysis.risks.length === 0) {
        event.completed({ allowEvent: true });
        return;
      }

      event.completed({ allowEvent: false, errorMessage: buildAlertMessage(analysis) });
    } catch (error) {
      // Never leave the send hanging: always complete.
      try {
        event.completed({
          allowEvent: false,
          errorMessage: "Recipient Guard could not complete its check. Review your recipients before sending."
        });
      } catch (ignored) { /* nothing more we can do */ }
    }
  }

  // Register both ways: associate is required by the new Outlook JS-only runtime;
  // the global name is the fallback for the HTML function-file path.
  if (typeof Office !== "undefined" && Office.actions && typeof Office.actions.associate === "function") {
    Office.actions.associate("onMessageSendDiagnostic", onMessageSendDiagnostic);
  }
  if (typeof globalThis !== "undefined") {
    globalThis.onMessageSendDiagnostic = onMessageSendDiagnostic;
  }
})();
