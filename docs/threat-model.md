# OLHelper threat model

## Scope and objective

This model covers the unified-manifest Outlook add-in, its Office WebView
runtime, NAA/MSAL authentication, direct Microsoft Graph calls, native Exchange
Inbox rules, Azure Static Web Apps hosting, the source/deployment pipeline, and
tenant administration. It supports review of the delegated `Mail.ReadWrite` and
`MailboxSettings.ReadWrite` request.

It does not model Outlook, Microsoft Entra ID, Microsoft Graph, Exchange Online,
GitHub, or Azure platform internals beyond the trust assumptions OLHelper relies
on.

## Architecture and data flow

```text
Approved user
    |
    | selects a message and confirms an operation
    v
Outlook + Office.js
    |
    | selected item ID and subject
    v
OLHelper WebView <------- static assets ------- Azure Static Web Apps
    |
    | NAA request
    v
Microsoft identity platform / broker
    |
    | delegated access token returned to WebView memory
    v
OLHelper central Graph client
    |
    | HTTPS + bearer token; allowlisted /me endpoints only
    v
Microsoft Graph
    |
    +---- Exchange folders and messages
    |
    +---- Exchange Inbox rules

Source repository -> protected CI/release workflow -> immutable release output
Tenant administrators -> consent, assignment, Conditional Access, revocation
```

The Azure Static Web Apps data path ends when static assets are delivered. Graph
tokens and Graph responses travel between the WebView and Microsoft services;
they are not proxied through the static host.

## Threat Modeling Portal submission alignment

The authoritative review record must be created in Threat Model Copilot (TMC)
and published to the Threat Modeling Portal. This Markdown document is the
design input, reviewer aid, and durable repository evidence; it does not replace
the portal-native model or the answers recorded in TMC.

Prepare the submission as follows:

1. Ensure the OLHelper Service Tree entry and accountable roles are current.
   Current TMC guidance permits Security Champion, Dev Owner, PM Owner, or
   Service Admin roles to initiate a review.
2. In [Threat Model Copilot](https://ai.security.azure/vnext), create a new
   iteration with a detailed title, description, and the applicable review team.
3. Create the following scenarios, each with its own data-flow diagram:
   - **Interactive mailbox operations:** user confirmation, Outlook/Office.js,
     WebView, NAA, Graph, Exchange folders/messages, and Inbox rules.
   - **Identity and administration:** Entra registration, delegated consent,
     enterprise-app assignment, Conditional Access, ownership, revocation, and
     incident containment.
   - **Software delivery:** repository, dependency resolution, CI validation,
     protected pilot environment, Static Web Apps, unified app package, and
     Outlook loading the approved origin.
4. Reproduce the components, external interactors, data flows, protocols,
   assets, and trust boundaries in this document on the TMC canvas. Draw
   executable-content delivery separately from mailbox-data flow and show that
   OLHelper has no application data store or backend.
5. Build the diagrams directly in TMC or import a `.tm7` file from the Microsoft
   Threat Modeling Tool. The portal/TMC history is the source of truth for the
   submitted model.
6. Run analysis and answer every generated question. Skip only inapplicable
   questions and record a specific justification in TMC.
7. Satisfy the TMC validation gates and use **Notify Review Team**. The security
   review team completes the asynchronous review and publishes accepted results
   and generated work items to the Threat Modeling Portal and Azure DevOps.
8. Triage every published work item, close inapplicable findings only with an
   auditable justification, and remediate applicable findings under the SDL Bug
   Bar. Repeat the review for the triggers below and any applicable recurring
   S360 cadence.

Portal scenario descriptions should link or attach this document, the reviewed
architecture diagram, permission-operation matrix, release evidence, tenant
control evidence, and the risk decision. Responses must be entered in TMC rather
than supplied only through email, Word, or this repository document.

## Assets

| Asset | Security objective |
| --- | --- |
| Delegated Graph access token | Prevent disclosure, persistence, replay, or use outside intended Graph operations |
| Mailbox messages and metadata | Maintain confidentiality; mutate only with explicit user intent |
| Folder and Inbox-rule configuration | Prevent unauthorized creation, movement, replacement, or deletion |
| User-to-selected-item context | Bind confirmation and mutation to the same item |
| Entra registration and grants | Prevent unauthorized redirect, scope, owner, consent, or assignment changes |
| Hosted JavaScript and dependencies | Preserve integrity because code executes with delegated token access |
| Unified app package | Preserve identity, origin, command, permission, and runtime integrity |
| Release pipeline and deployment credential | Prevent unauthorized production code delivery |
| Security and release evidence | Preserve integrity and availability without recording mailbox data |

## Actors and assumptions

### Expected actors

- Approved signed-in support engineer.
- Named application, tenant, repository, deployment, and incident-response
  owners.
- Microsoft identity, Office, Graph, Exchange, GitHub Actions, and Azure hosting
  services.

### Adversarial actors

- A malicious website attempting to embed or impersonate OLHelper.
- An attacker who compromises a user, WebView, dependency, repository account,
  workflow, deployment credential, registration owner, or Azure role.
- A user who creates a rule with an OLHelper-like name to induce unsafe
  modification.
- An approved user who accidentally confirms against a stale selected item.

### Trust assumptions

- Outlook supplies the active item and emits supported `ItemChanged` events.
- NAA returns tokens for the configured single-tenant client and trusted broker
  origin.
- Microsoft Graph enforces delegated scopes and Exchange mailbox permissions.
- TLS and the supported host platforms are not compromised.
- Tenant and release controls listed in `docs/security.md` are implemented.
- Production never enables arbitrary script sources, unreviewed runtime code, or
  user-controlled deployment origins.

## Trust boundaries

| Boundary | Data crossing | Principal concern |
| --- | --- | --- |
| User to OLHelper UI | Requested action and confirmation | Spoofing, stale context, unintended mutation |
| Outlook/Office.js to WebView | Selected item ID and subject | Incorrect item binding or host compromise |
| WebView to identity broker | Client, tenant, origin, scopes | Token issuance to the wrong app or origin |
| Broker to WebView memory | Delegated access token | Token disclosure to active JavaScript |
| WebView to Graph | Token, IDs, subjects, mutations | Excess authority, path confusion, response leakage |
| Graph to Exchange | Message, folder, and rule operations | Partial failure and native-rule semantics |
| Static host to WebView | Executable assets | XSS and supply-chain compromise |
| Repository/CI to static host and package | Release artifact | Unauthorized or non-reproducible deployment |
| Tenant administration to Entra | Consent, assignment, policy, ownership | Privilege expansion or weak containment |

## Graph operation matrix

| Operation | Method and allowed path shape | Data requested or sent | Permission |
| --- | --- | --- | --- |
| List/create root folders | `GET/POST /me/mailFolders` | Folder IDs, names, child counts; new display name | `Mail.ReadWrite` |
| List/create child folders | `GET/POST /me/mailFolders/{id}/childFolders` | Folder IDs/names; new display name | `Mail.ReadWrite` |
| Move case folder | `POST /me/mailFolders/{id}/move` | Destination folder ID: the Active or Archived container, or the well-known `deleteditems` folder for Delete case and emails | `Mail.ReadWrite` |
| Scan recent Inbox | `GET /me/mailFolders/inbox/messages` | `id,subject` only; newest first; maximum 250 | `Mail.ReadWrite` |
| Move message | `POST /me/messages/{id}/move` | Destination folder ID | `Mail.ReadWrite` |
| List/create rules | `GET/POST /me/mailFolders/inbox/messageRules` | Summary fields when listing; managed rule definition when creating | `MailboxSettings.ReadWrite` |
| Inspect/update/delete rule | `GET/PATCH/DELETE /me/mailFolders/inbox/messageRules/{id}` | Conditions, exceptions, actions, state, or intended update | `MailboxSettings.ReadWrite` |

The central client rejects every other Graph path. This is defense in depth
against programming errors and confused-deputy behavior, not a reduction in
delegated permission authority. Because Graph resource IDs are opaque, the
client validates their path position and the complete request shape but cannot
derive folder ancestry or rule ownership from an ID alone; the folder and rule
modules enforce those semantic relationships.

## Alternative architecture and decision

| Risk dimension | Delegated NAA design | Exchange RBAC backend |
| --- | --- | --- |
| Server-enforced mailbox scope | No OLHelper-specific subset; authority follows the signed-in user | Explicit Exchange management scope |
| Token exposure | Delegated token is available to trusted WebView JavaScript | App token stays in the confidential service |
| Standing and offline authority | User-present, delegated operation | App identity can operate without user presence |
| Compromise impact | Active user and mailboxes that user can access | All mailboxes assigned to the app's RBAC scope |
| Identity binding | `/me` binds the target to the signed-in actor | Backend must authenticate the actor and reject client-selected mailbox substitution |
| System attack surface | Static origin, WebView, runtime dependencies | Those client risks plus API, service identity, authorization, hosting, monitoring, and backend data handling |

The selected design minimizes standing privilege, cross-user authority,
credentials, backend exposure, and confused-deputy logic for OLHelper's
interactive user-owned-mailbox scenario. Its principal tradeoff is that a
delegated token enters the browser trust boundary and Graph cannot enforce an
OLHelper-folder-only scope.

Exchange RBAC would be preferred if the approval requirement is a
server-enforced mailbox population boundary or removal of Graph tokens from the
WebView. That change requires a confidential service, authoritative
actor-to-mailbox mapping, `/users/{mailbox}` calls, privacy-safe service
auditing, and a separate threat review. This decision remains subordinate to an
internal policy mandate requiring Exchange RBAC.

## STRIDE analysis

| ID | Category | Threat and abuse path | Impact | Controls | Residual risk |
| --- | --- | --- | --- | --- | --- |
| TM-01 | Spoofing | A malicious site frames or imitates OLHelper to collect input or trigger auth | Credential or user-intent theft | Exact HTTPS origin, single-tenant registration, trusted broker redirect, manifest-controlled URLs, CSP `frame-ancestors`, no referrer | A compromised approved origin remains trusted |
| TM-02 | Spoofing | Registration or manifest is changed to another client/origin | Token issued to attacker-controlled code | Restricted owners, protected configuration, package review/validation, release evidence, assignment | Privileged administrator compromise |
| TM-03 | Spoofing | Pinned pane displays item A but acts on newly selected item B | Wrong case is changed | Confirmation snapshot contains Outlook ID, Graph ID, and subject; `ItemChanged` disables/reloads; pre-mutation revalidation; fail-closed initialization | Host event failure outside documented behavior |
| TM-04 | Tampering | User-created rule copies an OLHelper display name | Unrelated rule is retargeted or deleted | Full condition/exception/action/state fingerprint; exact ID in name; reject collisions and modified rules | Exchange may add future harmless defaults that require compatibility updates |
| TM-05 | Tampering | Malformed Tracking ID or folder name alters routing | Misrouting or unsafe resource selection | Canonical marker, 16–19 digits, numeric boundary, URL encoding, ambiguity rejection, tests | Native rule prefix overlap remains |
| TM-06 | Tampering | Pagination or caller supplies an alternate method, URL, query, header, or body | Bearer token sent to attacker or broader Graph operation | Validation before token acquisition; exact HTTPS Graph origin; method/path/query/body contract; expected field selections and filters; supported pagination only; no caller headers, credentials, fragments, or unrelated request options | Same-origin malicious replacement code can bypass application helpers |
| TM-07 | Tampering | Dependency, action, workflow, or deployment token is compromised | Malicious JavaScript reaches every active user | Lockfile install, audit gates, Dependabot, dependency review, CodeQL, immutable action SHAs, protected environment, restricted owners, release hashes | Upstream zero-day or authorized malicious change |
| TM-08 | Repudiation | A mailbox mutation cannot be attributed | Weak incident reconstruction | User confirmation, tenant sign-in/audit logs, Graph/Exchange audit capabilities, release/workflow records | No application telemetry by design; evidence depends on tenant retention |
| TM-09 | Information disclosure | Token persists in browser storage, logs, an idle authenticated task pane, or a completed command runtime | Mailbox access until expiry/revocation | MSAL memory-only cache; no application persistence; prohibited sensitive logging; generic raw-error suppression; read-only checks retain the interactive session; completed mutations navigate after a visible ten-second countdown to an Office-initialized page that loads no MSAL or mailbox application code; short-lifetime command runtime releases its MSAL reference before `event.completed()` | Broker-controlled state, token availability between a status check and subsequent user action, active-operation memory exposure, host-controlled command-runtime destruction, and no cryptographic process-memory zeroization |
| TM-10 | Information disclosure | XSS or malicious hosted JavaScript reads tokens/Graph data | Mailbox disclosure within delegated scopes | Strict CSP, no inline/eval script requirement, constrained network destinations, reviewed origin, dependency controls | CSP permits Graph and identity because the app requires them; malicious same-origin code can use those destinations |
| TM-11 | Information disclosure | Overbroad queries retrieve unrelated mailbox content | Unnecessary privacy exposure | No body/attachment/recipient fields; `id,subject` Inbox selection; 250 cap; rule summary then candidate detail | Subjects of nonmatching recent Inbox messages are processed transiently |
| TM-12 | Information disclosure | Raw Graph or broker errors appear in UI/support artifacts | Tokens, identifiers, or diagnostic context disclosed | Graph client retains only HTTP status; long or token/broker-like errors are replaced; logging prohibition | New short sensitive non-Graph error formats may evade heuristics |
| TM-13 | Denial of service | Graph throttling, quota, large mailbox, or malformed state blocks workflow | Case routing unavailable or incomplete | Bounded scan, error surfacing, retry guidance, idempotent discovery, Repair action | No automatic throttling backoff or offline operation |
| TM-14 | Denial of service | Native rule quota is exhausted | Persistent routing cannot be created | Graph error surfaced without success-shaped fallback; existing state preserved where possible | User/admin remediation is required |
| TM-15 | Elevation of privilege | Unassigned user launches app or consent expands audience | Restricted mailbox capability reaches unintended users | Enterprise app requires assignment, controlled admin consent, restricted user consent, CA, periodic grant review | Incorrect group membership or policy exception |
| TM-16 | Elevation of privilege | Stolen token is used for mailbox APIs OLHelper never calls | Broader mailbox access than product behavior | Short-lived delegated token, CA/session controls, memory-only cache, bounded authenticated-document lifetime, incident revocation | OAuth scope is still `Mail.ReadWrite`; client allowlist does not bind a stolen token |
| TM-17 | Elevation of privilege | App-only credentials are introduced into browser code | Tenant-scale mailbox access | Architecture prohibits secrets, certificates, managed identity tokens, app permissions, and `/users` endpoints in the frontend | Future architecture changes require a new threat review |
| TM-18 | Integrity/availability | Multi-step Graph operation fails midway, the host terminates the runtime, or rollback fails | Folder/rule/message state diverges | Safe sequencing; session-end control disabled during active operations; limited rollback; precise partial-success reporting; status and Repair | Graph has no cross-resource transaction; Outlook or endpoint termination remains possible; rollback is not guaranteed |
| TM-19 | Tampering | A user deletes the wrong case, or case mail is lost, through Delete case and emails | Case folder and its emails are removed from the Active tree | Delete is a move to the well-known Deleted Items folder, never a permanent delete: the client allowlist rejects folder `DELETE` and `permanentDelete`. It offers only for an Active case folder identified by a status check, requires the same item-bound confirmation as other mutations, disables routing before the move and removes the managed rule after it, and re-enables routing if the move fails. Only the case folder is moved; Inbox, Sent Items, and other folders are untouched | Deleted Items can be emptied by the user or retention policy, after which recovery depends on Exchange Recoverable Items retention |

## Priority abuse cases

### Compromised static origin

An attacker who can deploy JavaScript to the approved origin can execute inside
the same trust boundary as OLHelper, request or observe delegated tokens, and
call Graph within granted scopes. Endpoint validation in OLHelper code cannot
constrain attacker-supplied replacement code. Preventive emphasis is therefore
on restricted Azure/repository roles, protected release approval, immutable
workflow dependencies, CSP, dependency governance, and rapid assignment/grant
revocation.

### Stolen delegated token

A stolen token can be replayed according to its claims until expiry or
revocation behavior takes effect. The token is not cryptographically restricted
to OLHelper folders, Tracking IDs, or endpoint allowlists. Memory-only caching,
no logging, Conditional Access, approved-user assignment, and emergency
revocation reduce likelihood or duration; they do not eliminate the impact.

### Conflicting Inbox rule

A rule name is user-controlled and is not proof of ownership. OLHelper retrieves
full details only for an exact-name candidate and validates its semantic
fingerprint. Any unexpected condition, exception, action, target, read-only
state, or reported error causes a fail-closed result. The user must resolve the
collision rather than OLHelper editing it.

### Selection changes after confirmation

Outlook permits a task pane to remain pinned while the selected item changes.
OLHelper binds confirmation to three item properties, registers `ItemChanged`,
disables actions on change, reloads the pane, and verifies the snapshot again
immediately before mutation.

### Tracking ID prefix overlap

Client-side matching rejects a digit following the validated Tracking ID.
Exchange native `subjectContains` cannot express that numeric boundary. A rule
for a shorter ID may therefore route mail for a longer ID. This is a platform
limitation and an explicit residual risk; a complete fix requires a
non-overlapping identifier format or a different routing service.

## Privacy and retention

Mailbox metadata exists only transiently in active JavaScript objects and Graph
transport. No mailbox content is intentionally written to the static host,
browser persistence, repository, CI artifacts, or application telemetry.
Release evidence contains code/package hashes and operational metadata only.

Support personnel must request sanitized reproduction steps and timestamps, not
message subjects, Tracking IDs, raw IDs, tokens, or raw auth/Graph responses.

## Verification evidence

Approval evidence should include:

- Source revision and successful CI workflow.
- Type-check, unit-test, dependency-audit, build, and manifest-validation results.
- Generated unified package and SHA-256 release list.
- Entra permission, consent, assignment, owner, and redirect-URI screenshots or
  exported configuration with personal data removed.
- Azure/repository role and protected-environment configuration.
- Production response headers and CSP.
- Negative test evidence for path allowlisting, item changes, rule collisions,
  malformed identifiers, partial failure, and raw-error suppression.
- Documented rollback, deployment-token rotation, grant/session revocation, and
  managed-rule cleanup exercise.

## Review triggers

Repeat the threat review before:

- Adding any Graph scope, app-only identity, backend, telemetry, database, or
  non-Microsoft network destination.
- Reading bodies, attachments, recipients, sender details, or more mailbox
  history.
- Changing the hosting origin, authentication library/flow, CSP, manifest
  authorization, command runtime, or mailbox target model.
- Introducing third-party runtime scripts or changing the deployment trust
  model.
- Expanding assignment beyond the approved population.

## Risk decision

Production approval must explicitly accept the residual breadth of delegated
`Mail.ReadWrite`, active-origin/token compromise risk, native-rule prefix
overlap, bounded Inbox history, and nontransactional Graph behavior. If those
risks are unacceptable, the appropriate response is an architectural change
such as a dedicated mailbox or a confidential Exchange-RBAC service with
independently approved mailbox scoping—not a claim that browser-side checks
narrow OAuth. For the current interactive scope, delegated NAA is selected
because it has lower standing privilege and aggregate service complexity; the
decision does not supersede an internal policy requirement to use Exchange
RBAC.
