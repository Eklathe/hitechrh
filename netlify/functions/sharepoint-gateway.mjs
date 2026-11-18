import { createRemoteJWKSet, jwtVerify } from "jose";
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const VALID_DOC_ID = /^(?:0[1-9]|1[01])-(?:0|[1-9][0-9]{0,2})$/;
const ALLOWED_EXTENSIONS = new Set(["pdf","doc","docx","odt","xls","xlsx","csv","xml","txt","jpg","jpeg","png"]);
const MAX_BYTES = 50 * 1024 * 1024;
const DOCUMENT_COUNTS = Object.freeze({"01":11,"02":13,"03":12,"04":12,"05":12,"06":8,"07":12,"08":11,"09":8,"10":10,"11":9});
const GRAPH = "https://graph.microsoft.com/v1.0";

function json(status, data) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store, private",
      "X-Content-Type-Options": "nosniff",
      "Referrer-Policy": "no-referrer",
    },
  });
}

function getEnv() {
  const tenantId = process.env.ENTRA_TENANT_ID?.trim();
  const spaClientId = process.env.ENTRA_SPA_CLIENT_ID?.trim();
  const apiClientId = process.env.ENTRA_API_CLIENT_ID?.trim();
  const apiClientSecret = process.env.ENTRA_API_CLIENT_SECRET?.trim();
  const mapJSON = process.env.CLIENT_FOLDERS_JSON?.trim();
  const enabled = process.env.ENABLE_SHAREPOINT_UPLOADS === "true";
  if (![tenantId, spaClientId, apiClientId].every(v => GUID.test(v ?? ""))) return null;
  if (!enabled || !apiClientSecret || !mapJSON) return null;
  let folders;
  try {
    folders = JSON.parse(mapJSON);
    if (!folders || typeof folders !== "object" || Array.isArray(folders)) return null;
  } catch {
    return null;
  }
  return { tenantId, spaClientId, apiClientId, apiClientSecret, folders };
}

function publicConfig(settings, request) {
  if (!settings) {
    // Diagnostic limité aux Deploy Previews : jamais de valeurs, noms d'utilisateurs,
    // chemins SharePoint, identifiants de dossiers ou secrets dans la réponse.
    const out = { enabled: false };
    const isPreview = /^deploy-preview-[0-9]+--hitechrhfr\.netlify\.app$/i.test(new URL(request.url).hostname);
    if (isPreview) {
      const invalid = [];
      for (const key of ["ENTRA_TENANT_ID", "ENTRA_SPA_CLIENT_ID", "ENTRA_API_CLIENT_ID"]) {
        if (!GUID.test((process.env[key] ?? "").trim())) invalid.push(key);
      }
      if (!(process.env.ENTRA_API_CLIENT_SECRET ?? "").trim()) invalid.push("ENTRA_API_CLIENT_SECRET");
      const foldersRaw = (process.env.CLIENT_FOLDERS_JSON ?? "").trim();
      let foldersIssue = null;
      if (!foldersRaw) {
        foldersIssue = "missing_or_empty";
      } else {
        try {
          const map = JSON.parse(foldersRaw);
          if (typeof map === "string") foldersIssue = "json_string_instead_of_object";
          else if (!map || typeof map !== "object" || Array.isArray(map)) foldersIssue = "wrong_json_structure";
        } catch {
          foldersIssue = "invalid_json_syntax";
        }
      }
      if (foldersIssue) invalid.push("CLIENT_FOLDERS_JSON");
      if (process.env.ENABLE_SHAREPOINT_UPLOADS !== "true") invalid.push("ENABLE_SHAREPOINT_UPLOADS");
      out.previewChecks = {
        invalidVariables: invalid,
        context: "Deploy Previews",
        clientFoldersIssue: foldersIssue,
        diagnosticVersion: 2
      };
    }
    return json(200, out);
  }
  return json(200, {
    enabled: true,
    tenantId: settings.tenantId,
    spaClientId: settings.spaClientId,
    scope: "api://" + settings.apiClientId + "/upload",
  });
}

async function authorize(request, settings) {
  const match = /^Bearer ([A-Za-z0-9_.~-]+)$/.exec(request.headers.get("authorization") ?? "");
  if (!match) return null;
  const issuer = "https://login.microsoftonline.com/" + settings.tenantId + "/v2.0";
  const keys = createRemoteJWKSet(new URL("https://login.microsoftonline.com/" + settings.tenantId + "/discovery/v2.0/keys"));
  try {
    const { payload } = await jwtVerify(match[1], keys, {
      issuer,
      audience: [settings.apiClientId, "api://" + settings.apiClientId],
      algorithms: ["RS256"],
      clockTolerance: 30,
    });
    // La signature, l'audience, le scope, le tenant et l'application SPA sont tous vérifiés.
    // Le lien au dossier client provient exclusivement de l'oid dans notre tenant.
    if (
      payload.tid?.toLowerCase() !== settings.tenantId.toLowerCase() ||
      payload.ver !== "2.0" ||
      payload.azp?.toLowerCase() !== settings.spaClientId.toLowerCase() ||
      typeof payload.scp !== "string" ||
      !payload.scp.split(" ").includes("upload") ||
      !GUID.test(payload.oid ?? "")
    ) return null;
    const client = settings.folders[payload.oid.toLowerCase()];
    if (!client || typeof client !== "object" || Array.isArray(client)) return null;
    if (typeof client.driveId !== "string" || typeof client.folderId !== "string") return null;
    if (!/^[A-Za-z0-9._!~-]{8,200}$/.test(client.driveId)) return null;
    if (!/^[A-Za-z0-9._!~-]{2,200}$/.test(client.folderId)) return null;
    if (client.type !== "general" && client.type !== "btp") return null;
    return { client, oid: payload.oid };
  } catch {
    return null;
  }
}

function safeFileName(originalName) {
  if (typeof originalName !== "string") return null;
  const base = originalName.trim().split(/[/\\]/).pop() ?? "";
  const match = /\.([A-Za-z0-9]{1,6})$/.exec(base);
  if (!match || !ALLOWED_EXTENSIONS.has(match[1].toLowerCase())) return null;
  const ext = match[1].toLowerCase();
  const name = base.slice(0, -match[0].length)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 64) || "document";
  return name + "." + ext;
}

async function graphToken(settings) {
  const body = new URLSearchParams({
    client_id: settings.apiClientId,
    client_secret: settings.apiClientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const resp = await fetch("https://login.microsoftonline.com/" + settings.tenantId + "/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
    signal: AbortSignal.timeout(14000),
  });
  if (!resp.ok) throw new Error("token_http_" + resp.status);
  const result = await resp.json();
  if (typeof result.access_token !== "string") throw new Error("token_missing");
  return result.access_token;
}

async function createUploadSession(settings, client, documentId, originalName, size) {
  const normalized = safeFileName(originalName);
  if (!normalized) return { error: "Type de fichier non autorisé.", status: 422 };
  if (!Number.isSafeInteger(size) || size < 1 || size > MAX_BYTES)
    return { error: "Taille de fichier non autorisée (maximum 50 Mo).", status: 422 };
  if (typeof documentId !== "string" || !VALID_DOC_ID.test(documentId))
    return { error: "Référence documentaire invalide.", status: 422 };
  const [category, docIndex] = documentId.split("-");
  if (!(category in DOCUMENT_COUNTS) || Number(docIndex) >= DOCUMENT_COUNTS[category])
    return { error: "Cette pièce ne figure pas dans la liste autorisée.", status: 422 };
  if (client.type !== "btp" && Number(category) >= 9)
    return { error: "Cette rubrique est réservée au parcours BTP.", status: 403 };

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const storedName = documentId.replace("-", "_") + "_" + stamp + "_" + crypto.randomUUID().slice(0, 8) + "_" + normalized;
  const token = await graphToken(settings);
  const path = GRAPH + "/drives/" + encodeURIComponent(client.driveId) + "/items/" +
    encodeURIComponent(client.folderId) + ":/" + encodeURIComponent(storedName) + ":/createUploadSession";
  const resp = await fetch(path, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename", name: storedName } }),
    signal: AbortSignal.timeout(15000),
  });
  if (!resp.ok) {
    // Ne jamais journaliser un jeton ni un nom de document client.
    console.error("SharePoint upload-session refused; HTTP", resp.status);
    return { error: "Le dossier SharePoint n'est pas disponible ou l'application ne dispose pas des droits nécessaires.", status: 502 };
  }
  const result = await resp.json();
  if (typeof result.uploadUrl !== "string" || !result.uploadUrl.startsWith("https://"))
    return { error: "La session de dépôt SharePoint n'a pas pu être ouverte.", status: 502 };

  return {
    uploadUrl: result.uploadUrl,
    expiresAt: result.expirationDateTime ?? null,
    storedName,
    maxBytes: MAX_BYTES,
  };
}

export default async function sharepointGateway(request) {
  if (!["GET", "POST"].includes(request.method)) return json(405, { error: "Méthode non autorisée." });
  const settings = getEnv();
  if (request.method === "GET") return publicConfig(settings, request);
  if (!settings) return json(503, { error: "Le dépôt sécurisé n'est pas encore activé." });
  const origin = request.headers.get("origin");
  if (origin && origin !== new URL(request.url).origin) return json(403, { error: "Origine refusée." });
  if (!(request.headers.get("content-type") ?? "").toLowerCase().startsWith("application/json"))
    return json(415, { error: "Format de requête invalide." });
  const bodyText = await request.text();
  if (bodyText.length > 8192) return json(413, { error: "Requête trop volumineuse." });

  // Contrôle d'accès obligatoire avant toute création de session Microsoft Graph.
  const access = await authorize(request, settings);
  if (!access) return json(403, { error: "Accès non autorisé ou aucun dossier client attribué." });
  let body;
  try { body = JSON.parse(bodyText); } catch { return json(400, { error: "JSON invalide." }); }
  if (!body || typeof body !== "object") return json(400, { error: "Requête invalide." });

  if (body.action === "profile") {
    return json(200, {
      client: String(access.client.name || "Votre entreprise").slice(0, 120),
      type: access.client.type,
      maxBytes: MAX_BYTES,
    });
  }
  if (body.action !== "createUploadSession") return json(400, { error: "Action inconnue." });
  try {
    const result = await createUploadSession(settings, access.client, body.documentId, body.fileName, body.size);
    if (result.error) return json(result.status, { error: result.error });
    return json(200, result);
  } catch {
    console.error("SharePoint upload-session failed");
    return json(502, { error: "Impossible de contacter Microsoft 365. Veuillez réessayer ultérieurement." });
  }
}
