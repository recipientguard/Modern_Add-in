// Centered review dialog, opened by the task pane via displayDialogAsync.
//
// A dialog runs in its own sandbox with NO access to the mailbox item, so every
// action here is posted to the parent pane (messageParent). The pane performs the
// send / whitelist against the item and messages results back (messageChild).
// If the parent can't deliver the payload (older Dialog API), the pane's own
// inline review panel is the fallback and this dialog shows a short notice.
(function () {
  "use strict";

  // Findings + the full recipient list, cached so the delay view can list
  // everyone and Cancel can restore the review view.
  var state = { items: [], recipients: [] };
  var countdownTimer = null;

  function post(obj) {
    try { Office.context.ui.messageParent(JSON.stringify(obj)); } catch (e) { /* ignore */ }
  }

  Office.onReady(function () {
    document.getElementById("dlgSend").addEventListener("click", function () { post({ action: "send" }); disableActions(); status("Sending…"); });
    document.getElementById("dlgDelay").addEventListener("click", showDelayConfirmation);
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
    if (msg.type === "payload") {
      state.items = msg.items || [];
      state.recipients = msg.recipients || [];
      renderReview();
    } else if (msg.type === "whitelisted") {
      markWhitelisted(msg.email);
    }
  }

  // --- review view ---

  function renderReview() {
    document.getElementById("dlgTitle").textContent = "Review before sending";
    document.getElementById("dlgIntro").textContent = "Recipient Guard paused a send. Please check the recipients are correct.";
    document.getElementById("dlgActions").hidden = false;
    status("");

    var body = document.getElementById("dlgBody");
    body.innerHTML = "";
    body.className = "";
    body.setAttribute("data-loaded", "1");

    var summary = document.createElement("div");
    summary.className = "check-summary warning";
    summary.textContent = state.items.length === 1
      ? "1 recipient to review before sending"
      : state.items.length + " recipients to review before sending";
    body.appendChild(summary);

    var list = document.createElement("div");
    list.className = "recipient-list";
    state.items.forEach(function (it) {
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

    // Full recipient list, so the user sees everyone this is going to.
    var sentHeading = document.createElement("div");
    sentHeading.className = "list-heading";
    sentHeading.textContent = "This message will be sent to:";
    body.appendChild(sentHeading);
    body.appendChild(buildRecipientList(state.recipients));
  }

  // Shared recipient list (email + To/Cc/Bcc label), used by the review view
  // and the delay-send confirmation so both list recipients identically.
  function buildRecipientList(recipients) {
    var list = document.createElement("div");
    list.className = "recipient-list";
    if (!recipients || recipients.length === 0) {
      var none = document.createElement("div");
      none.className = "muted";
      none.textContent = "(Recipient list unavailable.)";
      list.appendChild(none);
      return list;
    }
    recipients.forEach(function (rcpt) {
      var row = document.createElement("div");
      row.className = "recipient";
      var addr = document.createElement("strong");
      addr.textContent = rcpt.email;
      row.appendChild(addr);
      if (rcpt.type) {
        var badge = document.createElement("span");
        badge.className = "badge-muted";
        badge.textContent = rcpt.type;
        row.appendChild(badge);
      }
      list.appendChild(row);
    });
    return list;
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

  // --- delay-send confirmation view ---
  // Lists ALL recipients clearly with a countdown and a Cancel button. On
  // completion it posts the send; Cancel returns to the review view.

  function showDelayConfirmation() {
    document.getElementById("dlgTitle").textContent = "Sending in 60 seconds";
    document.getElementById("dlgIntro").textContent = "This message will be sent to the recipients below. Cancel if anything looks wrong.";
    document.getElementById("dlgActions").hidden = true;

    var body = document.getElementById("dlgBody");
    body.innerHTML = "";
    body.className = "";
    body.appendChild(buildRecipientList(state.recipients));

    var remaining = 60;
    var statusEl = document.getElementById("dlgStatus");
    statusEl.innerHTML = "";
    var label = document.createElement("span");
    label.textContent = "Sending in " + remaining + "s… ";
    var cancel = document.createElement("button");
    cancel.type = "button";
    cancel.textContent = "Cancel send";
    statusEl.appendChild(label);
    statusEl.appendChild(cancel);

    countdownTimer = setInterval(function () {
      remaining -= 1;
      if (remaining <= 0) {
        clearInterval(countdownTimer);
        countdownTimer = null;
        statusEl.textContent = "Sending…";
        post({ action: "send" });
        return;
      }
      label.textContent = "Sending in " + remaining + "s… ";
    }, 1000);

    cancel.addEventListener("click", function () {
      if (countdownTimer) { clearInterval(countdownTimer); countdownTimer = null; }
      renderReview();
    });
  }

  // --- helpers ---

  function disableActions() {
    ["dlgSend", "dlgDelay"].forEach(function (id) {
      var b = document.getElementById(id);
      if (b) b.disabled = true;
    });
  }

  function status(text) {
    document.getElementById("dlgStatus").textContent = text;
  }
})();
