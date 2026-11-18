import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

const gateway = "/.netlify/functions/sharepoint-gateway";
const connectButton = document.getElementById("client-connect");
const disconnectButton = document.getElementById("client-disconnect");
const title = document.getElementById("client-auth-title");
const info = document.getElementById("client-auth-description");
const clientName = document.getElementById("client-auth-company");
const uploadPanel = document.getElementById("client-upload-panel");
const uploadPicker = document.getElementById("client-upload-file");
const uploadCategory = document.getElementById("client-upload-category");
const uploadDocLabel = document.getElementById("client-upload-doc-label");
const uploadButton = document.getElementById("client-upload-confirm");
const uploadStatus = document.getElementById("client-upload-feedback");

let msal;
let settings;
let account;
let selectedDocument;
let authorizedProfile;

function status(message, error = false) {
  uploadStatus.textContent = message;
  uploadStatus.className = error ? "feedback error" : "feedback success";
}
function explain(message) {
  info.textContent = message;
}
function showPicker(docId, label) {
  if (!account || !authorizedProfile) return;
  selectedDocument = docId;
  uploadCategory.textContent = docId;
  uploadDocLabel.textContent = label;
  uploadPicker.value = "";
  uploadStatus.textContent = "";
  uploadStatus.className = "feedback";
  uploadPanel.hidden = false;
  uploadPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });
  uploadPicker.focus();
}

function updateConnected(connected) {
  connectButton.hidden = connected;
  disconnectButton.hidden = !connected;
  clientName.hidden = !connected;
  if (!connected) {
    clientName.textContent = "";
    authorizedProfile = null;
    uploadPanel.hidden = true;
    document.body.dataset.sharepointConnected = "false";
  } else {
    document.body.dataset.sharepointConnected = "true";
  }
}
async function getToken() {
  const request = { scopes: [settings.scope], account };
  try {
    const result = await msal.acquireTokenSilent(request);
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const result = await msal.acquireTokenPopup(request);
      return result.accessToken;
    }
    throw error;
  }
}
async function api(action, fields = {}) {
  const accessToken = await getToken();
  const response = await fetch(gateway, {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ action, ...fields }),
    cache: "no-store",
    credentials: "same-origin",
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(result.error || "La connexion au service de dépôt a échoué.");
  return result;
}
async function authenticate(force = false) {
  try {
    if (force) {
      const result = await msal.loginPopup({ scopes: [settings.scope], prompt: "select_account" });
      account = result.account;
    } else {
      account = msal.getAllAccounts()[0] || null;
      if (!account) return;
    }
    title.textContent = "Vérification de votre dossier client…";
    const profile = await api("profile");
    authorizedProfile = profile;
    clientName.textContent = "Dossier SharePoint : " + profile.client;
    title.textContent = "Espace documentaire connecté";
    explain("Vous pouvez maintenant déposer vos pièces dans le dossier SharePoint attribué à votre entreprise.");
    updateConnected(true);
    // Le parcours est déterminé par le rattachement serveur et non par le client.
    const mode = profile.type === "btp" ? "btp" : "general";
    document.querySelector('[data-mode="' + mode + '"]')?.click();
    document.dispatchEvent(new Event("hitechrh:render-list"));
    document.querySelectorAll("[data-mode]").forEach(button => {
      button.disabled = true;
      button.title = "Parcours défini par votre dossier client";
    });
  } catch (error) {
    account = null;
    updateConnected(false);
    title.textContent = "Accès au dossier client non autorisé";
    explain(error.message || "Le compte Microsoft n'est pas encore rattaché à un dossier de collecte. Contactez HI-TECH RH.");
  }
}

async function uploadInChunks(uploadUrl, file) {
  // Les fragments Microsoft Graph doivent être des multiples de 320 KiB, sauf le dernier.
  const chunkSize = 5 * 1024 * 1024;
  let start = 0;
  let complete = false;
  while (start < file.size) {
    const end = Math.min(start + chunkSize, file.size);
    const chunk = file.slice(start, end);
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + file.size,
        "Content-Type": "application/octet-stream",
      },
      body: chunk,
      // Le jeton Graph est encapsulé dans l'URL de session.
      // Aucun token Entra ni Authorization ne doit être joint à cette requête.
    });
    if (!res.ok) throw new Error("Échec d'envoi du fragment " + Math.ceil(end/chunkSize) + ". HTTP " + res.status);
    if (res.status === 200 || res.status === 201) {
      complete = true;
      start = end;
      break;
    }
    if (res.status !== 202) throw new Error("Réponse SharePoint inattendue (" + res.status + ").");
    start = end;
    status("Téléversement en cours : " + Math.round(100 * start / file.size) + " %");
  }
  if (!complete || start < file.size) throw new Error("SharePoint n'a pas confirmé l'enregistrement du document.");
}

connectButton.addEventListener("click", () => authenticate(true));
disconnectButton.addEventListener("click", async () => {
  updateConnected(false);
  account = null;
  document.querySelectorAll(".doc-upload-button").forEach(button => button.remove());
  document.querySelectorAll("[data-mode]").forEach(button => {button.disabled = false; button.title = "";});
  title.textContent = "Connexion à votre espace client";
  explain("Connectez-vous avec le compte Microsoft autorisé par HI-TECH RH.");
  if (msal) await msal.logoutPopup({ mainWindowRedirectUri: window.location.href });
});
uploadButton.addEventListener("click", async () => {
  const file = uploadPicker.files?.[0];
  if (!file || !selectedDocument) return status("Sélectionnez d'abord un fichier.", true);
  if (!account || !authorizedProfile) return status("Connectez-vous avant le dépôt.", true);
  if (file.size < 1 || file.size > 50*1024*1024) return status("Taille du fichier invalide : 50 Mo maximum.", true);
  const allowed = /\.(pdf|doc|docx|odt|xls|xlsx|csv|xml|txt|jpe?g|png)$/i;
  if (!allowed.test(file.name)) return status("Format non autorisé.", true);
  uploadButton.disabled = true;
  status("Création d'une session d'envoi sécurisée…");
  try {
    const session = await api("createUploadSession", {
      documentId: selectedDocument,
      fileName: file.name,
      size: file.size,
    });
    await uploadInChunks(session.uploadUrl, file);
    status("Document reçu et enregistré dans votre dossier SharePoint.");
    const select = document.querySelector('.doc-status[data-id="' + selectedDocument + '"]');
    if (select) {
      select.value = "transmis";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
    uploadPicker.value = "";
  } catch(error) {
    status(error.message || "Le dépôt n'a pas abouti.", true);
  } finally {
    uploadButton.disabled = false;
  }
});
document.addEventListener("click", event => {
  const trigger = event.target.closest("[data-upload-doc]");
  if (!trigger || !authorizedProfile) return;
  const id = trigger.getAttribute("data-upload-doc");
  const label = trigger.getAttribute("data-upload-label");
  showPicker(id, label || "Document");
});
document.addEventListener("hitechrh:render-list", () => {
  if (!authorizedProfile) return;
  document.querySelectorAll(".doc-line").forEach(row => {
    if (row.querySelector("[data-upload-doc]")) return;
    const id = row.dataset.doc;
    const label = row.querySelector(".doc-label")?.firstChild?.textContent?.trim() || "Document";
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn doc-upload-button";
    button.textContent = "Déposer";
    button.dataset.uploadDoc = id;
    button.dataset.uploadLabel = label;
    row.appendChild(button);
  });
});

async function initialize() {
  try {
    const response = await fetch(gateway, { cache: "no-store" });
    settings = await response.json();
    if (!settings.enabled || !settings.tenantId || !settings.spaClientId || !settings.scope) {
      title.textContent = "Dépôt SharePoint en attente d'activation";
      explain("Le suivi de collecte reste disponible. La connexion Microsoft sera activée après configuration par HI-TECH RH.");
      connectButton.hidden = true;
      return;
    }
    msal = new PublicClientApplication({
      auth: {
        clientId: settings.spaClientId,
        authority: "https://login.microsoftonline.com/" + settings.tenantId,
        redirectUri: window.location.origin + "/audit-rh-paie",
        postLogoutRedirectUri: window.location.origin + "/audit-rh-paie",
      },
      cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
    });
    await msal.initialize();
    await msal.handleRedirectPromise();
    connectButton.hidden = false;
    await authenticate(false);
  } catch (error) {
    title.textContent = "Service de dépôt temporairement indisponible";
    explain("Impossible de démarrer le module d'authentification. Le suivi de collecte et l'export restent disponibles.");
    connectButton.hidden = true;
  }
}
initialize();
