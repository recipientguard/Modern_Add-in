(function () {
  "use strict";

  function onMessageSendHandler(event) {
    event.completed({
      allowEvent: false,
      errorMessage: "Recipient Guard POC is running. This diagnostic version intentionally stops the send."
    });
  }

  globalThis.onMessageSendHandler = onMessageSendHandler;

  if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
})();
