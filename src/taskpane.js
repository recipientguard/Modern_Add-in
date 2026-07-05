const ANALYSIS_CACHE_KEY = "recipientGuard.latestComposeAnalysis.v1";

Office.onReady(() => {
  const button = document.getElementById("analyzeButton");
  button.addEventListener("click", renderAnalysis);
  renderMailbox();
  renderAnalysis();
  registerRecipientChangeHandler();
});

function renderMailbox() {
  document.getElementById("mailbox").textContent = window.RecipientGuardPoc.getMailboxEmail() || "Unknown";
  document.getElementById("internalDomain").textContent = window.RecipientGuardPoc.getInternalDomain() || "Unknown";
}

function cacheAnalysis(analysis) {
  try {
    localStorage.setItem(ANALYSIS_CACHE_KEY, JSON.stringify({
      createdAt: Date.now(),
      mailboxEmail: analysis.mailboxEmail,
      internalDomain: analysis.internalDomain,
      recipients: analysis.recipients,
      flagged: analysis.flagged,
      hasWarnings: analysis.hasWarnings
    }));
  } catch (error) {
  }
}

function clearCachedAnalysis() {
  try {
    localStorage.removeItem(ANALYSIS_CACHE_KEY);
  } catch (error) {
  }
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
    cacheAnalysis(analysis);

    if (analysis.recipients.length === 0) {
      results.textContent = "No resolved recipients found yet.";
      results.className = "muted";
      return;
    }

    const summary = document.createElement("div");
    summary.className = analysis.hasWarnings ? "check-summary warning" : "check-summary clear";
    summary.textContent = analysis.hasWarnings
      ? analysis.flagged.length + (analysis.flagged.length === 1 ? " external recipient found" : " external recipients found")
      : "No external recipients found";

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
    results.appendChild(list);
  } catch (error) {
    results.textContent = "Recipient Guard could not read recipients in this compose window.";
    results.className = "muted";
    clearCachedAnalysis();
  } finally {
    button.disabled = false;
    button.textContent = "Check recipients";
  }
}
