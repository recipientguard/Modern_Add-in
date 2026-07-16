// Nested App Authentication (NAA) + Microsoft Graph — A1 proof.
// Bundled by esbuild into src/naa.bundle.js (resolves the @azure/msal-browser
// import) and loaded by the task pane. Exposes window.RGNaa.
//
// Goal for A1: prove we can silently get a Graph token for the signed-in user
// (no backend) and read /me/people (frequently-contacted people = the modern
// equivalent of the Classic AutoComplete/Contacts known-identity source).

import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

// Public client ID from the Entra app registration (client IDs are not secrets).
// Registered in the RecipientGuard Ltd tenant (multitenant, SPA redirect
// brk-multihub://addin.recipientguard.co.uk).
const CLIENT_ID = "7519a415-3e8b-4c8e-9599-740a658ae7a2";
const GRAPH_SCOPES = ["People.Read", "User.Read"];

let msalInstance;

async function initMsal() {
  if (!msalInstance) {
    msalInstance = await createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/common"
      },
      cache: { cacheLocation: "localStorage" }
    });
  }
  return msalInstance;
}

function isNaaSupported() {
  try {
    return Boolean(Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
  } catch (e) {
    return false;
  }
}

// Silent first; fall back to an interactive popup only when the user hasn't yet
// consented (InteractionRequiredAuthError). Popups only work in the task pane —
// never in an event runtime.
async function acquireGraphToken() {
  const msal = await initMsal();
  const request = { scopes: GRAPH_SCOPES };
  try {
    const result = await msal.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const result = await msal.acquireTokenPopup(request);
      return result.accessToken;
    }
    throw err;
  }
}

// Fetch the signed-in user's relevant people (ranked by interaction). Returns a
// compact [{ displayName, email }] list.
async function getKnownPeople(top) {
  const token = await acquireGraphToken();
  const url =
    "https://graph.microsoft.com/v1.0/me/people?$top=" + (top || 25) +
    "&$select=displayName,scoredEmailAddresses";
  const response = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!response.ok) {
    const text = await response.text();
    throw new Error("Graph /me/people failed: " + response.status + " " + text);
  }
  const data = await response.json();
  return (data.value || [])
    .map(function (person) {
      const scored = person.scoredEmailAddresses || [];
      return {
        displayName: person.displayName || "",
        email: (scored[0] && scored[0].address) || ""
      };
    })
    .filter(function (person) { return person.email; });
}

window.RGNaa = {
  isNaaSupported: isNaaSupported,
  acquireGraphToken: acquireGraphToken,
  getKnownPeople: getKnownPeople
};
