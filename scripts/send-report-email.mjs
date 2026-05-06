import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { ClientSecretCredential } from "@azure/identity";

/** Bump when changing diagnostics so Actions logs prove which script ran. */
const SCRIPT_DIAG_VERSION = "2026-02-08c";

const REQUIRED_ENV_KEYS = [
  "AZURE_TENANT_ID",
  "AZURE_CLIENT_ID",
  "AZURE_CLIENT_SECRET",
  "OUTLOOK_SENDER_UPN",
  "OUTLOOK_TO",
];

function getMissingEnv() {
  return REQUIRED_ENV_KEYS.filter((key) => !process.env[key]?.trim());
}

function parseRecipients(raw) {
  return raw
    .split(/[;,]/)
    .map((value) => value.trim())
    .filter(Boolean);
}

function buildDefaultSubject() {
  const date = new Intl.DateTimeFormat("ja-JP", {
    timeZone: "Asia/Tokyo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
  return `WMS 第3週レポート (${date})`;
}

function resolveReportPath() {
  const fromEnv = process.env.REPORT_HTML_PATH?.trim() || "report-preview.html";
  if (path.isAbsolute(fromEnv)) {
    return fromEnv;
  }

  const thisDir = path.dirname(fileURLToPath(import.meta.url));
  return path.resolve(thisDir, "..", fromEnv);
}

async function fetchGraphAccessToken() {
  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID,
    process.env.AZURE_CLIENT_ID,
    process.env.AZURE_CLIENT_SECRET,
  );

  try {
    const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
    if (!tokenResponse?.token) {
      throw new Error("Token response had no token.");
    }
    return tokenResponse.token;
  } catch (err) {
    const msg = err?.message ?? String(err);
    throw new Error(
      `Failed to acquire Microsoft Graph access token (check AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET): ${msg}`,
    );
  }
}

function summarizeGraphFailure(response, responseBody) {
  const parts = [];
  parts.push(responseBody?.trim() ? responseBody.trim() : "body=(empty)");

  const wwwAuth = response.headers.get("www-authenticate");
  if (wwwAuth) {
    parts.push(`WWW-Authenticate: ${wwwAuth}`);
  }

  const reqId =
    response.headers.get("request-id") ??
    response.headers.get("x-ms-request-id") ??
    response.headers.get("client-request-id");
  if (reqId) {
    parts.push(`request-id: ${reqId}`);
  }

  return parts.join(" | ");
}

/** Decode JWT payload only (no verification) — for debugging tenant / roles. Never log the raw token. */
function decodeJwtPayload(token) {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) {
      return null;
    }
    const json = Buffer.from(parts[1].replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf8");
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function logTokenDiagnostics(token) {
  const p = decodeJwtPayload(token);
  if (!p) {
    console.error("Token diagnostics: could not decode JWT (unexpected token shape).");
    return;
  }
  const roles = Array.isArray(p.roles) ? p.roles : [];
  console.error(
    `Token diagnostics: tid=${p.tid ?? "?"} aud=${p.aud ?? "?"} appid=${p.appid ?? "?"} roles=${JSON.stringify(roles)}`,
  );
}

function logGraphFailureLines(status, responseBody, response) {
  console.error(`GRAPH_FAIL status=${status}`);
  console.error(`GRAPH_FAIL body=${responseBody?.trim() ? responseBody.trim() : "(empty)"}`);
  const wwwAuth = response.headers.get("www-authenticate");
  if (wwwAuth) {
    console.error(`GRAPH_FAIL WWW-Authenticate=${wwwAuth}`);
  }
  const reqId =
    response.headers.get("request-id") ??
    response.headers.get("x-ms-request-id") ??
    response.headers.get("client-request-id");
  if (reqId) {
    console.error(`GRAPH_FAIL request-id=${reqId}`);
  }
  try {
    for (const [key, value] of response.headers.entries()) {
      if (/^(authorization|set-cookie)$/i.test(key)) {
        continue;
      }
      console.error(`GRAPH_FAIL header ${key}: ${value}`);
    }
  } catch {
    /* ignore */
  }
}

async function verifySenderExistsInTenant(token, senderUpn) {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderUpn)}?$select=id,userPrincipalName,mail`;
  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const body = await r.text();
  const snippet = body.length > 800 ? `${body.slice(0, 800)}…` : body;
  console.error(`USER_LOOKUP status=${r.status} sender=${senderUpn}`);
  console.error(`USER_LOOKUP body=${snippet || "(empty)"}`);
  if (r.status === 403) {
    console.error(
      "USER_LOOKUP 403: Mail.Send alone cannot read user profiles. Add Microsoft Graph APPLICATION permission User.Read.All + Grant admin consent — OR ignore this check and verify sender manually in Entra ID → Users.",
    );
  } else if (r.status === 404) {
    console.error(
      "USER_LOOKUP 404: sender UPN not found in this tenant. Use Entra ID → Users → copy User principal name for OUTLOOK_SENDER_UPN.",
    );
  } else if (!r.ok) {
    console.error("USER_LOOKUP failed; sender may still be valid if Mail.Send is granted.");
  }
}

function toInternetMessageAttachment(htmlContent, reportFilePath) {
  const reportFileName = path.basename(reportFilePath);
  return {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: reportFileName,
    contentType: "text/html",
    contentBytes: Buffer.from(htmlContent, "utf8").toString("base64"),
  };
}

async function sendMail() {
  const missingEnv = getMissingEnv();
  if (missingEnv.length > 0) {
    throw new Error(`Missing required environment variables: ${missingEnv.join(", ")}`);
  }

  const reportFilePath = resolveReportPath();
  const htmlBody = await readFile(reportFilePath, "utf8");
  const recipients = parseRecipients(process.env.OUTLOOK_TO);

  if (recipients.length === 0) {
    throw new Error("OUTLOOK_TO does not contain any valid recipients.");
  }

  const subject = process.env.REPORT_SUBJECT?.trim() || buildDefaultSubject();
  const shouldAttachReport = process.env.ATTACH_REPORT !== "false";

  console.error(`send-report-email.mjs diagnostics=${SCRIPT_DIAG_VERSION}`);

  const token = await fetchGraphAccessToken();
  logTokenDiagnostics(token);

  await verifySenderExistsInTenant(token, process.env.OUTLOOK_SENDER_UPN.trim());

  const mailPayload = {
    message: {
      subject,
      body: {
        contentType: "HTML",
        content: htmlBody,
      },
      toRecipients: recipients.map((address) => ({
        emailAddress: { address },
      })),
      attachments: shouldAttachReport ? [toInternetMessageAttachment(htmlBody, reportFilePath)] : [],
    },
    saveToSentItems: true,
  };

  const sender = encodeURIComponent(process.env.OUTLOOK_SENDER_UPN);
  const endpoint = `https://graph.microsoft.com/v1.0/users/${sender}/sendMail`;

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(mailPayload),
  });

  if (!response.ok) {
    const responseBody = await response.text();
    logGraphFailureLines(response.status, responseBody, response);
    const detail = summarizeGraphFailure(response, responseBody);
    throw new Error(`Graph sendMail failed (${response.status}): ${detail}`);
  }

  console.log(`Mail accepted by Graph. subject="${subject}" to=${recipients.join(", ")}`);
}

sendMail().catch((error) => {
  console.error(error.message);
  if (error.stack) {
    console.error(error.stack);
  }
  process.exitCode = 1;
});
