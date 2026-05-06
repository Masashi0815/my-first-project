# Outlook Auto Report Setup (GitHub Actions + Node + Graph)

This project sends `report-preview.html` as an Outlook email using Microsoft Graph.

## 1) Azure app registration

1. Open Azure Portal -> Microsoft Entra ID -> App registrations -> New registration.
2. Save:
   - Application (client) ID
   - Directory (tenant) ID
3. Create a client secret under Certificates & secrets.
4. API permissions:
   - Microsoft Graph -> Application permissions -> `Mail.Send`
   - (Optional, for script user lookup only) also add `User.Read.All`
5. Click **Grant admin consent** for the tenant.

`Mail.Send` sends mail; it does **not** allow reading user directory (`GET /users`). Optional `User.Read.All` lets the diagnostic `USER_LOOKUP` step succeed instead of HTTP 403.

## 2) Mailbox settings

Decide sender and recipients:

- Sender mailbox UPN (example): `report-bot@your-company.com`
- Recipients list (comma separated): `a@example.com,b@example.com`

`OUTLOOK_SENDER_UPN` should be a mailbox the app is allowed to send from.

Optional (recommended for guest UPNs with `#` in the name): set **`OUTLOOK_SENDER_OBJECT_ID`** to the user’s **Object ID** (from Entra ID → user → Overview). The script then calls `.../users/{object-id}/sendMail` and avoids URL issues with special characters.

Optional: set **`SKIP_USER_LOOKUP`** to `true` to skip the diagnostic `GET /users` call (or add **User.Read.All** app permission if you want the lookup to succeed).

## 3) GitHub repository secrets

Set these repository secrets in GitHub:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `OUTLOOK_SENDER_UPN` (unless you use `OUTLOOK_SENDER_OBJECT_ID` only)
- `OUTLOOK_TO`
- (Optional) `OUTLOOK_SENDER_OBJECT_ID`, `SKIP_USER_LOOKUP`

## 4) Local test

From repository root:

```bash
npm install
npm run send:report
```

Optional env vars:

- `REPORT_HTML_PATH` (default: `report-preview.html`)
- `REPORT_SUBJECT` (default auto-generated with date)
- `ATTACH_REPORT` (`true` by default, set `false` to disable attachment)

## 5) GitHub Actions schedule

Workflow file: `.github/workflows/send-report.yml`

- Scheduled: every Friday 09:00 UTC (18:00 JST)
- Manual run: `workflow_dispatch`

You can adjust cron in the workflow if needed.
