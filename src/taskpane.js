Office.onReady(() => {
  const button = document.getElementById("analyzeButton");
  button.addEventListener("click", renderAnalysis);
  renderMailbox();
  renderReviewFromContext();
  renderAnalysis();
  renderWhitelist();
  registerRecipientChangeHandler();

  const peopleButton = document.getElementById("loadPeopleButton");
  if (peopleButton) peopleButton.addEventListener("click", loadKnownPeople);

  const reviewSendButton = document.getElementById("reviewSendButton");
  if (reviewSendButton) reviewSendButton.addEventListener("click", () => reviewSend(0));
  const reviewDelayButton = document.getElementById("reviewDelayButton");
  if (reviewDelayButton) reviewDelayButton.addEventListener("click", () => reviewSend(60));
});

// Send the message from the review pane. The Smart Alert dialog is gone by the
// time this pane is open, so this is how the user proceeds without re-triggering
// the block: set a one-shot bypass, then sendAsync. Optionally count down first.
// NOTE: the countdown runs in this pane — if the pane is closed mid-countdown the
// send won't fire (Modern has no background delayed-send like Classic did).
function reviewSend(delaySeconds) {
  const RG = window.RecipientGuardPoc;
  const item = Office.context.mailbox.item;
  const status = document.getElementById("reviewStatus");
  const sendBtn = document.getElementById("reviewSendButton");
  const delayBtn = document.getElementById("reviewDelayButton");

  if (!item || typeof item.sendAsync !== "function") {
    status.textContent = "This Outlook version can't send from here — go back to your message and choose Send Anyway.";
    return;
  }

  const reset = () => {
    if (sendBtn) sendBtn.disabled = false;
    if (delayBtn) delayBtn.disabled = false;
  };
  const doSend = () => {
    status.textContent = "Sending...";
    // Set the bypass, then send. Nothing after sendAsync is guaranteed to run
    // (the item may already be sent), so clear-on-failure is the only follow-up.
    RG.setSendBypass().then(() => {
      item.sendAsync((res) => {
        if (res && res.status === Office.AsyncResultStatus.Failed) {
          RG.clearSendBypass();
          status.textContent = "Couldn't send: " + ((res.error && res.error.message) || "unknown error");
          reset();
        }
      });
    });
  };

  if (sendBtn) sendBtn.disabled = true;
  if (delayBtn) delayBtn.disabled = true;

  if (!delaySeconds) {
    doSend();
    return;
  }

  let remaining = delaySeconds;
  status.innerHTML = "";
  const label = document.createElement("span");
  label.textContent = "Sending in " + remaining + "s… ";
  const cancel = document.createElement("button");
  cancel.type = "button";
  cancel.className = "linklike";
  cancel.textContent = "Cancel";
  status.appendChild(label);
  status.appendChild(cancel);

  const timer = setInterval(() => {
    remaining -= 1;
    if (remaining <= 0) {
      clearInterval(timer);
      status.textContent = "";
      doSend();
      return;
    }
    label.textContent = "Sending in " + remaining + "s… ";
  }, 1000);

  cancel.addEventListener("click", () => {
    clearInterval(timer);
    RG.clearSendBypass();
    status.textContent = "Delayed send cancelled.";
    reset();
  });
}

// When the pane is opened from the Smart Alert "review" button, Outlook hands us
// the findings as an initialization context. Render them prominently at the top.
// When the pane is opened normally there's no context, so this is a no-op.
function renderReviewFromContext() {
  const panel = document.getElementById("reviewPanel");
  const item = Office.context.mailbox.item;
  if (!panel || !item || typeof item.getInitializationContextAsync !== "function") return;

  item.getInitializationContextAsync((asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) return;
    const raw = asyncResult.value;
    if (!raw) return;

    let risks;
    try {
      risks = (JSON.parse(raw) || {}).risks || [];
    } catch (error) {
      return;
    }
    if (risks.length === 0) return;

    const body = document.getElementById("reviewBody");
    body.innerHTML = "";

    const summary = document.createElement("div");
    summary.className = "check-summary warning";
    summary.textContent = risks.length === 1
      ? "1 recipient to review before sending"
      : risks.length + " recipients to review before sending";
    body.appendChild(summary);

    const list = document.createElement("div");
    list.className = "recipient-list";
    risks.forEach((risk) => list.appendChild(buildRiskRow(risk)));
    body.appendChild(list);

    panel.hidden = false;
    panel.scrollIntoView({ block: "start" });

    // Present the review as a centered modal (closer to the Classic form). The
    // inline panel above stays as the fallback if the dialog can't open.
    openReviewDialog();
  });
}

// Open the centered review dialog and host its message channel. The dialog has
// no access to the mail item, so it posts actions back here and we act on them.
let reviewDialog = null;
function openReviewDialog() {
  const ui = Office.context.ui;
  if (!ui || typeof ui.displayDialogAsync !== "function") return; // inline panel is the fallback
  const url = window.location.origin + "/src/review-dialog.html?v=20260718-1";
  // One recipient-centric list: every recipient, flagged ones annotated inline.
  const recipientsPromise = window.RecipientGuardPoc.analyzeCurrentMessage()
    .then((a) => buildDialogRecipients(a.risks || [], a.recipients || []))
    .catch(() => []);

  ui.displayDialogAsync(url, { height: 52, width: 42, displayInIframe: true }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded || !res.value) return;
    const dialog = res.value;
    reviewDialog = dialog;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => handleDialogMessage(dialog, recipientsPromise, arg.message));
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { reviewDialog = null; });
  });
}

// Annotate every recipient with whether/why it's flagged, so the dialog shows a
// single combined list instead of a separate red findings card + a plain list.
function buildDialogRecipients(risks, recipients) {
  const RG = window.RecipientGuardPoc;
  const noteByEmail = Object.create(null);
  risks.forEach((risk) => {
    const note = RG.noteForRisk(risk);
    if (!note) return;
    (risk.emails || []).forEach((email) => {
      if (!noteByEmail[email]) noteByEmail[email] = note; // first (strongest) wins; condense() prevents overlap
    });
  });
  return recipients.map((r) => {
    const flagged = Boolean(noteByEmail[r.email]);
    const dom = RG.getDomain(r.email);
    return {
      email: r.email,
      type: r.type,
      flagged: flagged,
      note: noteByEmail[r.email] || null,
      whitelistEmail: flagged ? r.email : null,
      whitelistDomain: (flagged && dom && !RG.isPublicDomain(dom)) ? dom : null
    };
  });
}

function handleDialogMessage(dialog, recipientsPromise, raw) {
  let msg;
  try {
    msg = JSON.parse(raw);
  } catch (error) {
    return;
  }
  const RG = window.RecipientGuardPoc;

  if (msg.action === "ready") {
    // Hand the dialog the annotated recipient list. If messageChild isn't
    // supported, the dialog shows its "use the panel" notice and the inline
    // panel covers it.
    recipientsPromise.then((recipients) => {
      try {
        dialog.messageChild(JSON.stringify({ type: "payload", recipients: recipients }));
      } catch (error) {
        /* older Dialog API — fallback handled in the dialog */
      }
    });
  } else if (msg.action === "cancel") {
    safeCloseReviewDialog(dialog);
  } else if (msg.action === "whitelist") {
    RG.addToWhitelist(msg.email).then(() => {
      renderAnalysis();
      renderWhitelist();
      try {
        dialog.messageChild(JSON.stringify({ type: "whitelisted", email: msg.email }));
      } catch (error) {
        /* no ack channel — the pane list still refreshed */
      }
    });
  } else if (msg.action === "whitelistDomain") {
    RG.addDomainToWhitelist(msg.domain).then(() => {
      renderAnalysis();
      renderWhitelist();
      try {
        dialog.messageChild(JSON.stringify({ type: "whitelisted", domain: msg.domain }));
      } catch (error) {
        /* no ack channel — the pane list still refreshed */
      }
    });
  } else if (msg.action === "send") {
    // One-shot bypass so this send passes without re-prompting, then send.
    RG.setSendBypass().then(() => {
      const item = Office.context.mailbox.item;
      safeCloseReviewDialog(dialog);
      if (item && typeof item.sendAsync === "function") {
        item.sendAsync((r) => {
          if (r && r.status === Office.AsyncResultStatus.Failed) RG.clearSendBypass();
        });
      }
    });
  }
}

function safeCloseReviewDialog(dialog) {
  try {
    dialog.close();
  } catch (error) {
    /* already closed */
  }
  reviewDialog = null;
}

// One source of truth for how a finding reads: title, detail line, and (for the
// single-subject rules) which address a whitelist action would target. Group
// rules (same display name / username) span several addresses, so whitelisting
// would be ambiguous — no whitelistEmail. Used by the pane rows AND the dialog
// payload so both surfaces word findings identically.
function describeRisk(risk) {
  const RG = window.RecipientGuardPoc;
  let title;
  let detail = RG.noteForRisk(risk);
  let whitelistEmail = null;
  if (risk.ruleId === "known_alternative") {
    title = "Possibly wrong recipient: " + risk.emails[0];
    whitelistEmail = risk.emails[0];
  } else if (risk.ruleId === "same_display_name") {
    title = 'Shares a display name: "' + (risk.displayName || "").trim() + '"';
    detail = (risk.emails || []).join(", ");
  } else if (risk.ruleId === "same_localpart_different_domain") {
    title = 'Same username: "' + risk.localPart + '"';
    detail = (risk.emails || []).join(", ");
  } else {
    title = "External recipient: " + (risk.emails || [])[0];
    whitelistEmail = (risk.emails || [])[0];
  }
  // Offer a domain-whitelist only for a non-public domain — trusting the whole of
  // gmail.com etc. would defeat the tool, so those get address-only.
  let whitelistDomain = null;
  if (whitelistEmail) {
    const dom = RG.getDomain(whitelistEmail);
    if (dom && !RG.isPublicDomain(dom)) whitelistDomain = dom;
  }
  return { title, detail, whitelistEmail, whitelistDomain };
}

// Build one flagged-recipient row. Shared by the live "Check recipients" list
// and the review-before-sending panel so both explain a finding identically.
function buildRiskRow(risk) {
  const d = describeRisk(risk);
  const row = document.createElement("div");
  row.className = "recipient flagged";

  const title = document.createElement("strong");
  title.textContent = d.title;
  row.appendChild(title);

  const meta = document.createElement("div");
  meta.className = "muted";
  meta.textContent = d.detail;
  row.appendChild(meta);

  if (d.whitelistEmail) {
    row.appendChild(buildWhitelistActions(d.whitelistEmail, d.whitelistDomain, row, meta));
  }
  return row;
}

function linklike(text, onClick) {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "linklike";
  b.textContent = text;
  b.addEventListener("click", onClick);
  return b;
}

function markWhitelistedRow(row, meta, message) {
  row.className = "recipient whitelisted";
  meta.textContent = "Whitelisted — " + message;
  const actions = row.querySelector(".wl-actions");
  if (actions) actions.remove();
}

// Whitelist actions for a flagged row: always the safe per-ADDRESS option; the
// per-DOMAIN option only when the domain isn't public, and gated behind an inline
// "trust everyone at X?" confirm because it's a much broader action.
function buildWhitelistActions(email, domain, row, meta) {
  const RG = window.RecipientGuardPoc;
  const actions = document.createElement("div");
  actions.className = "wl-actions";

  function renderNormal() {
    actions.innerHTML = "";
    actions.appendChild(linklike("Don't warn about this address", async () => {
      await RG.addToWhitelist(email);
      markWhitelistedRow(row, meta, "won't warn about " + email + " again");
      renderAnalysis();
      renderWhitelist();
    }));
    if (domain) {
      actions.appendChild(linklike("Don't warn about anyone at " + domain, renderConfirm));
    }
  }

  function renderConfirm() {
    actions.innerHTML = "";
    const q = document.createElement("span");
    q.className = "muted";
    q.textContent = "Trust everyone at " + domain + "? ";
    actions.appendChild(q);
    actions.appendChild(linklike("Trust domain", async () => {
      await RG.addDomainToWhitelist(domain);
      markWhitelistedRow(row, meta, "won't warn about anyone at " + domain);
      renderAnalysis();
      renderWhitelist();
    }));
    actions.appendChild(linklike("Cancel", renderNormal));
  }

  renderNormal();
  return actions;
}

// The "Trusted addresses & domains" manage view — the ONLY place to remove a
// whitelist entry (there was previously no way to un-whitelist). Add is
// contextual-only (from a flagged recipient), so no free-text domain box exists
// that could trust an unseen or public domain.
function renderWhitelist() {
  const container = document.getElementById("whitelistManage");
  if (!container) return;
  const RG = window.RecipientGuardPoc;
  const wl = RG.readWhitelist();
  const entries = (wl.domains || []).map((d) => ({ value: d, label: "Anyone at " + d }))
    .concat((wl.emails || []).map((e) => ({ value: e, label: e })));

  container.innerHTML = "";
  if (entries.length === 0) {
    const none = document.createElement("div");
    none.className = "muted";
    none.textContent = "Nothing trusted yet. Use “Don't warn…” on a flagged recipient to add one.";
    container.appendChild(none);
    return;
  }

  entries.forEach((entry) => {
    const row = document.createElement("div");
    row.className = "wl-row";
    const label = document.createElement("span");
    label.textContent = entry.label;
    row.appendChild(label);
    row.appendChild(linklike("Remove", async () => {
      await RG.removeFromWhitelist(entry.value);
      renderWhitelist();
      renderAnalysis();
    }));
    container.appendChild(row);
  });
}

// Load the MSAL/NAA bundle on demand (it's ~652 KB, so we keep it off the
// pane-open path and fetch it only when the user turns on smart detection).
let naaLoadPromise = null;
function ensureNaaLoaded() {
  if (window.RGNaa) return Promise.resolve();
  if (naaLoadPromise) return naaLoadPromise;
  naaLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "./naa.bundle.js";
    script.onload = () => resolve();
    script.onerror = () => {
      naaLoadPromise = null; // allow a retry on the next click
      reject(new Error("Failed to load naa.bundle.js"));
    };
    document.head.appendChild(script);
  });
  return naaLoadPromise;
}

// A1 proof: acquire a Graph token via NAA and list the user's frequently-
// contacted people. This is the known-identity source we'll compare against.
async function loadKnownPeople() {
  const out = document.getElementById("peopleResults");
  const button = document.getElementById("loadPeopleButton");

  out.className = "muted";
  out.textContent = "Preparing sign-in…";
  try {
    await ensureNaaLoaded();
  } catch (error) {
    out.textContent = "Couldn't load the sign-in module. Check your connection and try again.";
    return;
  }
  if (!window.RGNaa) {
    out.textContent = "Auth module not loaded.";
    return;
  }
  if (!window.RGNaa.isNaaSupported()) {
    out.textContent = "This Outlook version doesn't support the sign-in method needed (NAA). Try new Outlook or Outlook on the web.";
    return;
  }

  button.disabled = true;
  button.textContent = "Loading...";
  out.textContent = "Getting your contacts (you may be asked to grant permission)...";
  out.className = "muted";

  try {
    const people = await window.RGNaa.getKnownPeople(50);
    if (people.length === 0) {
      out.textContent = "Connected, but no relevant people were returned.";
      return;
    }

    // Cache a compact known-identity list to RoamingSettings so the send-event
    // runtime can read it (cross-runtime; localStorage can't do this).
    const RG = window.RecipientGuardPoc;
    const records = people.map(RG.toKnownRecord);
    await RG.writeKnownIdentities(records);

    // Now that known identities are cached, re-run the recipient analysis so the
    // "Current recipients" section reflects them immediately.
    renderAnalysis();

    out.innerHTML = "";
    out.className = "";
    const heading = document.createElement("div");
    heading.className = "check-summary clear";
    heading.textContent = "Smart detection on — " + people.length + " known contacts loaded";
    out.appendChild(heading);

    const list = document.createElement("div");
    list.className = "recipient-list";
    people.slice(0, 15).forEach((person) => {
      const row = document.createElement("div");
      row.className = "recipient";
      const title = document.createElement("strong");
      title.textContent = person.displayName || person.email;
      row.appendChild(title);
      const meta = document.createElement("div");
      meta.className = "muted";
      meta.textContent = person.email;
      row.appendChild(meta);
      list.appendChild(row);
    });
    out.appendChild(list);
  } catch (error) {
    out.textContent = "Couldn't load contacts: " + (error && error.message ? error.message : String(error));
    out.className = "muted";
  } finally {
    button.disabled = false;
    button.textContent = "Refresh contacts";
  }
}

function renderMailbox() {
  document.getElementById("mailbox").textContent = window.RecipientGuardPoc.getMailboxEmail() || "Unknown";
  document.getElementById("internalDomain").textContent = window.RecipientGuardPoc.getInternalDomain() || "Unknown";
}

function registerRecipientChangeHandler() {
  const item = Office.context.mailbox.item;
  if (!item || typeof item.addHandlerAsync !== "function" || !Office.EventType || !Office.EventType.RecipientsChanged) {
    return;
  }

  item.addHandlerAsync(Office.EventType.RecipientsChanged, () => {
    renderAnalysis();
  });
}

async function renderAnalysis() {
  const results = document.getElementById("results");
  const button = document.getElementById("analyzeButton");
  button.disabled = true;
  button.textContent = "Checking...";
  results.textContent = "Checking recipients...";
  results.className = "muted";

  try {
    const analysis = await window.RecipientGuardPoc.analyzeCurrentMessage();

    if (analysis.recipients.length === 0) {
      results.textContent = "No resolved recipients found yet.";
      results.className = "muted";
      return;
    }

    const risks = analysis.risks || [];
    const summary = document.createElement("div");
    summary.className = analysis.hasWarnings ? "check-summary warning" : "check-summary clear";
    summary.textContent = analysis.hasWarnings
      ? risks.length + (risks.length === 1 ? " potential issue found" : " potential issues found")
      : "No recipient issues found";

    const warnings = document.createElement("div");
    warnings.className = "recipient-list";
    risks.forEach((risk) => warnings.appendChild(buildRiskRow(risk)));

    const list = document.createElement("div");
    list.className = "recipient-list";

    analysis.recipients.forEach((recipient) => {
      const row = document.createElement("div");
      row.className = "recipient" + (recipient.isExternal ? " external" : "");

      const title = document.createElement("strong");
      title.textContent = recipient.email;
      row.appendChild(title);

      if (recipient.isExternal) {
        const badge = document.createElement("span");
        badge.className = "badge";
        badge.textContent = "External";
        row.appendChild(badge);
      }

      const meta = document.createElement("div");
      meta.className = "muted";
      meta.textContent = recipient.type + (recipient.name ? " - " + recipient.name : "");
      row.appendChild(meta);

      list.appendChild(row);
    });

    results.innerHTML = "";
    results.className = "";
    results.appendChild(summary);
    if (risks.length > 0) {
      results.appendChild(warnings);
    }
    results.appendChild(list);
  } catch (error) {
    results.textContent = "Recipient Guard could not read recipients in this compose window.";
    results.className = "muted";
  } finally {
    button.disabled = false;
    button.textContent = "Check recipients";
  }
}
