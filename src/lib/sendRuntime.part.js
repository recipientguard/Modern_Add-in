// Send-event handler (OnMessageSend). Combined with engine.js by the build
// into src/sendTestRuntime.js. Reads the recipients inline at send time and
// always calls event.completed — a send must never be left hanging.

var SEND_SAFETY_TIMEOUT_MS = 3000; // fail-open well under Outlook's 5s limit

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
    var knownIdentities = readKnownIdentities();
    getAllRecipients().then(function (recipients) {
      clearTimeout(safety);
      var risks = computeRisks(recipients, internalDomain, knownIdentities);
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

function registerSendHandler() {
  try {
    Office.actions.associate("onMessageSendDiagnostic", onMessageSendDiagnostic);
  } catch (e) { /* associate may not be ready yet; onReady path covers it */ }
}

// IMPORTANT: on new Outlook / Outlook on the web, Office.actions.associate only
// binds the handler when called inside Office.onReady(). Registering only at
// top level (as the yo-office template does) succeeds silently but Outlook
// never dispatches the event. So we register in BOTH places.
registerSendHandler();
if (typeof globalThis !== "undefined") {
  globalThis.onMessageSendDiagnostic = onMessageSendDiagnostic;
}
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(registerSendHandler);
}
