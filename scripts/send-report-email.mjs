import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { ClientSecretCredential } from "@azure/identity";

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

  const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
  if (!tokenResponse?.token) {
    throw new Error("Failed to acquire Microsoft Graph access token.");
  }
  return tokenResponse.token;
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
  const token = await fetchGraphAccessToken();

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
    throw new Error(`Graph sendMail failed (${response.status}): ${responseBody}`);
  }

  console.log(`Mail accepted by Graph. subject="${subject}" to=${recipients.join(", ")}`);
}

sendMail().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
