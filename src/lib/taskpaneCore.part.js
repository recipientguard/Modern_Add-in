// Task-pane API surface. Combined with engine.js by the build into
// src/recipientGuardCore.js. Exposes window.RecipientGuardPoc, consumed by
// src/taskpane.js.

function analyzeCurrentMessage() {
  var mailboxEmail = getMailboxEmail();
  var internalDomain = getInternalDomain();
  var whitelist = readWhitelist();
  return getAllRecipients().then(function (recipients) {
    // Annotate explicitly for the task-pane's per-recipient "External" badges —
    // deliberately not a side effect of detection, so rendering never depends
    // on which rules ran or in what order.
    recipients.forEach(function (recipient) {
      recipient.isExternal = isExternalRecipient(recipient, internalDomain);
      recipient.isWhitelisted = isWhitelisted(recipient.email, whitelist);
    });

    var risks = computeRisks(recipients, internalDomain, readKnownIdentities(), whitelist);
    return {
      mailboxEmail: mailboxEmail,
      internalDomain: internalDomain,
      recipients: recipients,
      risks: risks,
      hasWarnings: risks.length > 0
    };
  });
}

var globalScope = typeof window !== "undefined" ? window : globalThis;
globalScope.RecipientGuardPoc = {
  analyzeCurrentMessage: analyzeCurrentMessage,
  buildAlertMessage: buildAlertMessage,
  computeRisks: computeRisks,
  getInternalDomain: getInternalDomain,
  getMailboxEmail: getMailboxEmail,
  // known-identity store (Graph -> RoamingSettings)
  toKnownRecord: toKnownRecord,
  readKnownIdentities: readKnownIdentities,
  writeKnownIdentities: writeKnownIdentities,
  // whitelist (per-address "don't warn again")
  readWhitelist: readWhitelist,
  isWhitelisted: isWhitelisted,
  addToWhitelist: addToWhitelist,
  // one-shot send bypass (pane "send now")
  setSendBypass: setSendBypass,
  clearSendBypass: clearSendBypass,
  // exposed for reuse/testing
  normalizeName: normalizeName,
  getLocalPart: getLocalPart,
  getDomain: getDomain
};
