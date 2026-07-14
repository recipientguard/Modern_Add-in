// Centered review dialog, opened by the task pane via displayDialogAsync.
//
// A dialog runs in its own sandbox with NO access to the mailbox item, so every
// action here is posted to the parent pane (messageParent). The pane performs the
// send / whitelist against the item and messages results back (messageChild).
// If the parent can't deliver the payload (older Dialog API), the pane's own
// inline review panel is the fallback and this dialog shows a short notice.
(function () {
  "use strict";

  function post(obj) {
    try { Office.context.ui.messageParent(JSON.stringify(obj)); } catch (e) { /* ignore */ }
  }

  Office.onReady(function () {
    document.getElementById("dlgSend").addEventListener("click", function () { startSend(0); });
    document.getElementById("dlgDelay").addEventListener("click", function () { startSend(60); });
    document.getElementById("dlgCancel").addEventListener("click", function () { post({ action: "cancel" }); });

    try {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage);
    } catch (e) { /* DialogApi 1.2 unavailable — the load timeout below covers it */ }

    post({ action: "ready" });

    // If the payload never arrives, point the user at the side panel instead.
    setTimeout(function () {
      var body = document.getElementById("dlgBody");
      if (body && body.getAttribute("data-loaded") !== "1") {
        body.textContent = "Couldn't load the details here — use the Recipient Guard panel to review these recipients.";
      }
    }, 2500);
  });

  function onParentMessage(arg) {
    var msg;
    try { msg = JSON.parse(arg.message); } catch (e) { return; }
    if (msg.type === "payload") render(msg.items || []);
    else if (msg.type === "whitelisted") markWhitelisted(msg.email);
  }

  function render(items) {
    var body = document.getElementById("dlgBody");
    body.innerHTML = "";
    body.className = "";
    body.setAttribute("data-loaded", "1");

    var summary = document.createElement("div");
    summary.className = "check-summary warning";
    summary.textContent = items.length === 1
      ? "1 recipient to review before sending"
      : items.length + " recipients to review before sending";
    body.appendChild(summary);

    var list = document.createElement("div");
    list.className = "recipient-list";
    items.forEach(function (it) {
      var row = document.createElement("div");
      row.className = "recipient external";
      row.setAttribute("data-email", it.whitelistEmail || "");

      var title = document.createElement("strong");
      title.textContent = it.title;
      row.appendChild(title);

      if (it.detail) {
        var meta = document.createElement("div");
        meta.className = "muted";
        meta.textContent = it.detail;
        row.appendChild(meta);
      }

      if (it.whitelistEmail) {
        var wl = document.createElement("button");
        wl.type = "button";
        wl.className = "linklike";
        wl.textContent = "Don't warn about this address";
        wl.addEventListener("click", function () {
          wl.disabled = true;
          post({ action: "whitelist", email: it.whitelistEmail });
        });
        row.appendChild(wl);
      }
      list.appendChild(row);
    });
    body.appendChild(list);
  }

  function markWhitelisted(email) {
    var rows = document.querySelectorAll(".recipient[data-email]");
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].getAttribute("data-email") !== email) continue;
      rows[i].className = "recipient whitelisted";
      var btn = rows[i].querySelector("button.linklike");
      if (btn) btn.parentNode.removeChild(btn);
      var note = rows[i].querySelector(".muted");
      if (note) note.textContent = "Whitelisted — won't warn about " + email + " again";
    }
  }

  var countdownTimer = null;
  function startSend(delaySeconds) {
    var status = document.getElementById("dlgStatus");
    setSendButtonsDisabled(true);

    if (!delaySeconds) {
      status.textContent = "Sending…";
      post({ action: "send" });
      return;
    }

    var remaining = delaySeconds;
    status.innerHTML = "";
    var label = document.createElement("span");
    label.textContent = "Sending in " + remaining + "s… ";
    var cancel = document.createElement("button");
    cancel.type = "button";
    cancel.className = "linklike";
    cancel.textContent = "Cancel";
    status.appendChild(label);
    status.appendChild(cancel);

    countdownTimer = setInterval(function () {
      remaining -= 1;
      if (remaining <= 0) {
        clearInterval(countdownTimer);
        status.textContent = "Sending…";
        post({ action: "send" });
        return;
      }
      label.textContent = "Sending in " + remaining + "s… ";
    }, 1000);

    cancel.addEventListener("click", function () {
      clearInterval(countdownTimer);
      status.textContent = "";
      setSendButtonsDisabled(false);
    });
  }

  function setSendButtonsDisabled(disabled) {
    ["dlgSend", "dlgDelay"].forEach(function (id) {
      var b = document.getElementById(id);
      if (b) b.disabled = disabled;
    });
  }
})();
