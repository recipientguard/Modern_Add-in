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

  // A pane-initiated "send now" set a one-shot bypass — the pane already showed
  // the review, so let this single send through without re-prompting.
  if (consumeSendBypass()) { finish({ allowEvent: true }); return; }

  // Never let the send hang: if analysis stalls, fail open (allow the send).
  var safety = setTimeout(function () { finish({ allowEvent: true }); }, SEND_SAFETY_TIMEOUT_MS);

  try {
    var internalDomain = getInternalDomain();
    var knownIdentities = readKnownIdentities();
    var whitelist = readWhitelist();
    getAllRecipients().then(function (recipients) {
      clearTimeout(safety);
      var risks = computeRisks(recipients, internalDomain, knownIdentities, whitelist);
      if (risks.length === 0) {
        finish({ allowEvent: true });
      } else {
        // Block, and offer to open the task pane for the full review list.
        // commandId points at the manifest's task-pane button; contextData
        // carries the findings the pane re-hydrates.
        finish({
          allowEvent: false,
          errorMessage: buildAlertMessage(risks),
          commandId: PANE_COMMAND_ID,
          contextData: buildContextData(risks)
        });
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
