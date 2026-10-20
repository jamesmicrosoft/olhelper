# OLHelper

OLHelper is an Outlook task-pane add-in that uses delegated Microsoft Graph
permissions to manage support-case folders, native Inbox rules, and selected
messages. The pilot supports a complete case lifecycle:

- **Track** creates or reuses an active case folder, moves the selected message,
  enables persistent routing, and sweeps matching messages from the 250 most
  recent Inbox messages.
- **Check status** reports whether the case is active, archived, untracked, or
  requires routing repair.
- **Archive** disables routing, moves the case folder to `Archived`, and removes
  the OLHelper-managed rule.
- **Delete case and emails** (active cases only) disables routing, moves the case
  folder and every email in it to Deleted Items instead of archiving, and removes
  the OLHelper-managed rule. It is a move, not a permanent delete, so the folder
  can be restored from Deleted Items until that folder is emptied. Mail in the
  Inbox, Sent Items, or any other folder is not touched.
- **Reopen** moves an archived case back to `Active` and restores routing.
- **Repair routing** recreates or enables the managed rule for an active case.

The pilot has no backend service, client secret, application-level mailbox
access, or centralized storage of mailbox content.

Case folders may be renamed with space-delimited context after the Tracking ID,
for example `1234567890123456 - Contoso`. OLHelper rejects ambiguous matches
rather than choosing between multiple folders for the same Tracking ID.

## Configuration responsibilities

- `package.json` declares the browser/runtime dependencies and the commands used
  to build, test, validate, audit, and package the add-in.
- `appPackage/manifest.json` is the unified Microsoft 365 manifest that tells
  Outlook when and where to display OLHelper. It is a template whose placeholder
  origin and Entra client ID are replaced during hosted package generation.
- `webpack.config.js` compiles the TypeScript and CSS in `src/`, creates the
  task-pane and command HTML pages, copies icons, and injects Entra identifiers.

## Validation commands

```powershell
npm install
npm run typecheck
npm test
npm run build
npm run validate
npm run security:audit
npm run security:audit:all
```

The generated files are written to `dist/`. Do not edit that directory.

## Hosted package generation

The checked-in manifest is a deployment template with placeholder origin and
Entra client ID values. Package generation replaces every template URL with the
repository-pinned Azure Static Web Apps origin. The protected deployment
environment supplies `OLHELPER_CLIENT_ID` and `OLHELPER_TENANT_ID`. Changing the
executable origin therefore requires a reviewed source change rather than an
environment-variable update. The Entra registration must include this
trusted-broker redirect:

```text
brk-multihub://<production-origin>
```

The broker redirect contains only the origin, without a path.

Local serving, certificate generation, and automated Outlook sideloading are
intentionally not part of the maintained toolchain. Test changes through the
hosted pilot deployment.

## Azure Hosted Pilot

The Azure-hosted pilot runs in an Azure subscription associated with a
different tenant from the M365 developer sandbox. Azure hosts the static files,
the sandbox app registration controls NAA identity and delegated Graph access.

See:

- [Azure Static Web Apps pilot deployment](docs/deploy-azure-static-web-apps.md)
- [Security architecture and release gates](docs/security.md)
- [Threat model and Graph permission justification](docs/threat-model.md)
- [Production readiness checklist](docs/production-readiness-checklist.md)

The deployment workflow generates
`appPackage/build/olhelper-pilot.zip` for the protected HTTPS origin. The ZIP
contains the generated `manifest.json` and required 32×32 outline and 192×192
color icons; executable add-in assets remain in Azure Static Web Apps.
