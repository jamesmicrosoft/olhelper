# OLHelper

OLHelper is an Outlook task-pane add-in that uses delegated Microsoft Graph
permissions to manage support-case folders, native Inbox rules, and selected
messages. The pilot supports a complete case lifecycle:

- **Track** creates or reuses an active case folder, moves the selected message,
  and enables persistent routing.
- **Check status** runs automatically when the task pane opens and reports whether
  the case is active, archived, untracked, or requires routing repair. Only the
  actions that apply to that state are shown. **Refresh status** re-runs the check.
- **Archive** disables routing, moves the case folder to `Archived`, and removes
  the OLHelper-managed rule.
- **Delete case and emails** (active cases only) disables routing, moves the case
  folder and every email in it to Deleted Items instead of archiving, and removes
  the OLHelper-managed rule. It asks for confirmation first and leaves mail in
  the Inbox and Sent Items alone.
- **Reopen** moves an archived case back to `Active` and restores routing.
- **Repair routing** recreates or enables the managed rule for an active case.

The pilot has no backend service, client secret, application-level mailbox
access, or centralized storage of mailbox content.

## Configuration responsibilities

- `package.json` declares the browser/runtime dependencies and the commands used
  to build, validate, serve, and sideload the add-in.
- `manifest.xml` tells Outlook when and where to display OLHelper. All local
  URLs point to the Webpack HTTPS server at `https://localhost:3000`. Production
  and pilot deployments utilize a central Static Web App for hosting.
- `webpack.config.js` compiles the TypeScript and CSS in `src/`, creates the
  task-pane and command HTML pages, copies icons, injects Entra identifiers, and
  serves the resulting files over trusted local HTTPS.
- `.env.local` contains the sandbox Entra application and tenant IDs used by
  Webpack. It is intentionally excluded from Git. These identifiers are not
  client secrets, but are excluded to ensure a fully PII-stripped public repo.

## Locally Hosted Sandbox Entra Application Testing

Register a single-tenant SPA in the Microsoft 365 developer sandbox:

1. Add the SPA redirect URI `brk-multihub://localhost:3000`.
2. Add delegated Microsoft Graph permissions:
   - `Mail.ReadWrite`
   - `MailboxSettings.ReadWrite`
3. Copy `.env.example` to `.env.local` and replace both placeholder IDs.

## Local commands

```powershell
npm install
npm run typecheck
npm run build
npm run validate
npm start
```

`npm start` starts the HTTPS development server and attempts to sideload
`manifest.xml` into Outlook. Use `npm run stop` to stop the debugging session.
To run only the local server for manual sideloading, use:

```powershell
npm run dev-server
```

The generated files are written to `dist/`. Do not edit that directory.

## Local and production URLs

The checked-in manifest is intentionally configured only for the local pilot.
For organizational deployment, every `https://localhost:3000` URL
is replaced with an approved static hosting origin with a production 
trusted-broker redirect URI to the Entra application:

```text
brk-multihub://<production-origin>
```

The broker redirect contains only the origin, without a path.

## Azure Hosted Pilot

The Azure-hosted pilot runs in an Azure subscription associated with a
different tenant from the M365 developer sandbox. Azure hosts the static files,
the sandbox app registration controls NAA identity and delegated Graph access.

See:

- [Azure Static Web Apps pilot deployment](docs/deploy-azure-static-web-apps.md)
- [Pilot security plan and release gates](docs/security.md)

The deployment workflow generates `dist/manifest.pilot.xml` for the supplied
HTTPS origin. The checked-in `manifest.xml` remains localhost-only.
