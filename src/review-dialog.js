// Centered review dialog, opened by the task pane via displayDialogAsync.
//
// A dialog runs in its own sandbox with NO access to the mailbox item, so every
// action here is posted to the parent pane (messageParent). The pane performs the
// send / whitelist against the item and messages results back (messageChild).
// If the parent can't deliver the payload (older Dialog API), the pane's own
// inline review panel is the fallback and this dialog shows a short notice.
(function () {
  "use strict";

  // The full recipient list (each annotated with flagged/note/whitelist targets),
  // cached so the delay view can reuse it and Cancel can restore the review view.
  var state = { recipients: [] };
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
      state.recipients = msg.recipients || [];
      renderReview();
    } else if (msg.type === "whitelisted") {
      if (msg.domain) markWhitelistedByDomain(msg.domain);
      else if (msg.email) markWhitelistedByEmail(msg.email);
    }
  }

  // --- review view: one combined recipient list, flagged rows annotated ---

  function renderReview() {
    document.getElementById("dlgTitle").textContent = "Review before sending";
    document.getElementById("dlgIntro").textContent = "Recipient Guard paused a send. Please check the recipients are correct.";
    document.getElementById("dlgActions").hidden = false;
    status("");

    var body = document.getElementById("dlgBody");
    body.innerHTML = "";
    body.className = "";
    body.setAttribute("data-loaded", "1");

    var heading = document.createElement("div");
    heading.className = "list-heading";
    heading.textContent = "This message will be sent to:";
    body.appendChild(heading);
    body.appendChild(buildRecipientList(state.recipients, true));
  }

  // Shared recipient list. showFlags=true annotates flagged recipients with a
  // reason + whitelist actions; false renders a plain "who it's going to" list.
  function buildRecipientList(recipients, showFlags) {
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
      row.className = "recipient" + (showFlags && rcpt.flagged ? " flagged" : "");
      row.setAttribute("data-email", rcpt.email);

      var head = document.createElement("div");
      var addr = document.createElement("strong");
      addr.textContent = rcpt.email;
      head.appendChild(addr);
      if (rcpt.type) {
        var badge = document.createElement("span");
        badge.className = "badge-muted";
        badge.textContent = rcpt.type;
        head.appendChild(badge);
      }
      row.appendChild(head);

      if (showFlags && rcpt.flagged) {
        var note = document.createElement("div");
        note.className = "flag-note";
        note.textContent = rcpt.note || "Possibly the wrong recipient";
        row.appendChild(note);
        if (rcpt.whitelistEmail) {
          row.appendChild(buildWhitelistActions(rcpt.whitelistEmail, rcpt.whitelistDomain));
        }
      }
      list.appendChild(row);
    });
    return list;
  }

  function linklike(text, onClick) {
    var b = document.createElement("button");
    b.type = "button";
    b.className = "linklike";
    b.textContent = text;
    b.addEventListener("click", onClick);
    return b;
  }

  // Whitelist actions: always the safe per-address option; the per-domain option
  // only when the domain isn't public (parent filtered those out), gated behind an
  // inline "trust everyone at X?" confirm. Actions post to the parent; the parent's
  // ack (markWhitelisted*) clears the affected row(s).
  function buildWhitelistActions(email, domain) {
    var actions = document.createElement("div");
    actions.className = "wl-actions";

    function renderNormal() {
      actions.innerHTML = "";
      actions.appendChild(linklike("Don't warn about this address", function () {
        post({ action: "whitelist", email: email });
      }));
      if (domain) {
        actions.appendChild(linklike("Don't warn about anyone at " + domain, renderConfirm));
      }
    }

    function renderConfirm() {
      actions.innerHTML = "";
      var q = document.createElement("span");
      q.className = "muted";
      q.textContent = "Trust everyone at " + domain + "? ";
      actions.appendChild(q);
      actions.appendChild(linklike("Trust domain", function () {
        post({ action: "whitelistDomain", domain: domain });
      }));
      actions.appendChild(linklike("Cancel", renderNormal));
    }

    renderNormal();
    return actions;
  }

  function clearFlag(row) {
    row.className = "recipient";
    var note = row.querySelector(".flag-note");
    if (note) note.parentNode.removeChild(note);
    var actions = row.querySelector(".wl-actions");
    if (actions) actions.parentNode.removeChild(actions);
  }

  function markWhitelistedByEmail(email) {
    var rows = document.querySelectorAll(".recipient[data-email]");
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].getAttribute("data-email") === email) clearFlag(rows[i]);
    }
  }

  function markWhitelistedByDomain(domain) {
    var suffix = "@" + domain;
    var rows = document.querySelectorAll(".recipient[data-email]");
    for (var i = 0; i < rows.length; i++) {
      var e = rows[i].getAttribute("data-email") || "";
      if (e.slice(-suffix.length) === suffix) clearFlag(rows[i]);
    }
  }

  // --- delay-send confirmation view ---
  // Lists ALL recipients (plainly) with a countdown and a Cancel button. On
  // completion it posts the send; Cancel returns to the review view.

  function showDelayConfirmation() {
    document.getElementById("dlgTitle").textContent = "Sending in 60 seconds";
    document.getElementById("dlgIntro").textContent = "This message will be sent to the recipients below. Cancel if anything looks wrong.";
    document.getElementById("dlgActions").hidden = true;

    var body = document.getElementById("dlgBody");
    body.innerHTML = "";
    body.className = "";
    body.appendChild(buildRecipientList(state.recipients, false));

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
