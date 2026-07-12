Office.onReady(() => {
  const button = document.getElementById("analyzeButton");
  button.addEventListener("click", renderAnalysis);
  renderMailbox();
  renderAnalysis();
  registerRecipientChangeHandler();

  const peopleButton = document.getElementById("loadPeopleButton");
  if (peopleButton) peopleButton.addEventListener("click", loadKnownPeople);
});

// A1 proof: acquire a Graph token via NAA and list the user's frequently-
// contacted people. This is the known-identity source we'll compare against.
async function loadKnownPeople() {
  const out = document.getElementById("peopleResults");
  const button = document.getElementById("loadPeopleButton");

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
    const people = await window.RGNaa.getKnownPeople(25);
    if (people.length === 0) {
      out.textContent = "Connected, but no relevant people were returned.";
      return;
    }
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
    risks.forEach((risk) => {
      const row = document.createElement("div");
      row.className = "recipient external";
      const title = document.createElement("strong");
      if (risk.ruleId === "same_display_name") {
        title.textContent = 'Same name, different addresses: "' + (risk.displayName || "").trim() + '"';
      } else if (risk.ruleId === "same_localpart_different_domain") {
        title.textContent = 'Same prefix, different domains: "' + risk.localPart + '"';
      } else {
        title.textContent = "External recipient";
      }
      row.appendChild(title);

      const meta = document.createElement("div");
      meta.className = "muted";
      meta.textContent = (risk.emails || []).join(", ");
      row.appendChild(meta);
      warnings.appendChild(row);
    });

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
