# OLHelper Security and Architecture

## Security position

OLHelper is a static Outlook add-in. Outlook loads the application in an Office
WebView, Nested App Authentication (NAA) obtains delegated Microsoft Graph
access for the signed-in user, and the WebView calls Microsoft Graph directly.
Azure Static Web Apps serves HTML, JavaScript, CSS, and images only.

OLHelper has no application backend, database, client secret, managed identity,
app-only Graph permission, or centralized mailbox-data store. The static host
does not receive Graph tokens or Graph responses. Tokens and mailbox data do
enter the hosted JavaScript process, so integrity of the deployed origin and
every runtime dependency is the primary security boundary.


## Authorization and permission justification

| Permission | Type | Required operations | Why a narrower permission is insufficient |
| --- | --- | --- | --- |
| `MailboxItem.Read.User` | Office add-in resource-specific consent | Read the currently selected item's subject and identifier through Office.js | This permission is selected-item access only. It does not authorize Graph folder, message-move, or rule operations. |
| `Mail.ReadWrite` | Delegated Microsoft Graph | Discover and create the OLHelper folder tree; scan only IDs and subjects from up to 250 recent Inbox messages; move matching messages and case folders, including moving a case folder to Deleted Items | Graph folder creation and message/folder moves require `Mail.ReadWrite`. Graph has no delegated scope limited to a Tracking ID, selected message, or OLHelper folder tree. |
| `MailboxSettings.ReadWrite` | Delegated Microsoft Graph | List, inspect, create, enable, disable, retarget, and delete native Inbox rules | Graph exposes no narrower delegated rule-management permission. |

The Graph scopes are powerful. Client validation and endpoint restrictions
constrain intended OLHelper behavior, but do not reduce the authority encoded in
a stolen token. In particular, they are not a substitute for tenant consent,
assignment, Conditional Access, deployment integrity, or XSS prevention.

There is no documented delegated mailbox-RSC replacement for these operations.
Exchange Online RBAC for Applications is an app-only mailbox-scoping model. It
would require a confidential backend or managed identity, application
authentication, `/users/{mailbox}` Graph calls, and an authoritative check that
the interactive user is allowed to operate on that mailbox. The current design is
chosen for it's smaller and well-defined threat scope.

## Architecture alternative assessment

Internal documentation describes the preference for Resource Specific Consent via
Exhange RBAC roles. The architectural and security differences between this model and the model used within this tool are defined below:

| Security property | Delegated NAA and direct Graph | Exchange RBAC and confidential backend |
| --- | --- | --- |
| Mailbox population boundary | Signed-in user's delegated mailbox access; no OLHelper-specific server-side mailbox subset | Exchange-enforced scope to configured mailboxes |
| Token location | Short-lived delegated token enters approved WebView memory | App token remains in the backend |
| Standing privilege | Requires an approved user and interactive sign-in | Service identity can act without user presence until disabled |
| Compromise blast radius | Active user and mailboxes that user can access | Every mailbox in the application's RBAC scope |
| Actor-to-mailbox binding | Naturally derived from the signed-in user and `/me` | Backend must prevent a confused deputy and derive the target mailbox authoritatively |
| Added attack surface | Static executable origin and browser dependencies | Public API, service identity, credential/managed identity, authorization layer, hosting, monitoring, and data-handling boundary |
| Central policy and auditing | Relies primarily on Entra, Graph, Exchange, release, and user confirmation evidence | Can add centralized policy enforcement and application audit events |

For OLHelper's interactive, user-owned-mailbox workflow, the current design is
assessed as lower in aggregate implementation complexity, standing privilege,
cross-user authority, and confused-deputy risk, provided the origin, release,
assignment, consent, and Conditional Access controls in this document are
implemented. This is a contextual-based architecture decision rather than a claim that browser-held delegated tokens are intrinsically safer.


## Data inventory and minimization

OLHelper processes:

- The selected item's Outlook ID, converted Graph ID, and subject.
- A validated 16–19 digit Tracking ID extracted from the subject.
- Folder names and folder IDs needed to find or manage the case folder.
- Rule summaries, followed by full details only for an exact OLHelper-managed
  rule-name candidate.
- IDs and subjects only for at most the 250 most recent Inbox messages during a
  Track sweep.

OLHelper does not request message bodies, attachments, recipients, sender
details, contacts, calendar data, or mailbox-wide search results. It does not
persist mailbox data in cookies, browser storage, a database, or telemetry.

The following are prohibited in logs, telemetry, URLs, analytics, crash reports,
or support screenshots:

- Access, refresh, or ID tokens and authorization headers.
- Message subjects, bodies, attachments, or recipient/sender data.
- Message, folder, or rule IDs.
- Tracking IDs.
- User email addresses or mailbox identifiers.
- Raw Graph, MSAL, NAA, OneAuth, or broker error payloads.

The UI replaces long or token/broker-like exception text with a generic error.
Azure hosting logs may contain static resource paths, network metadata, and
user-agent information; they must not contain mailbox data.

## Token handling

- MSAL is configured with `BrowserCacheLocation.MemoryStorage`.
- OLHelper never copies a token into `localStorage`, `sessionStorage`,
  IndexedDB, cookies, URLs, application logs, or telemetry.
- A token is attached only by the central Graph client and only to an allowed
  `https://graph.microsoft.com/v1.0` request.
- Read-only status checks keep the task-pane session active so the user can
  proceed directly to Track, Archive, Delete, Reopen, or Repair. After a completed
  mailbox mutation, including a failed mutation, OLHelper disables further
  mailbox actions, displays an accessible ten-second visual countdown, and
  navigates to a minimal session-ended document. An explicit **End secure
  session** control performs the same navigation immediately. Session
  termination is disabled while a Graph workflow is active so it cannot
  interrupt a multi-request folder, rule, or message operation.
- The session-ended document initializes Office.js so Outlook recognizes it as
  a valid add-in page, but it does not load MSAL, Graph, case-workflow, or
  task-pane application code. It applies Office theme colors when available,
  with light, dark, forced-color, keyboard-focus, and reduced-motion fallbacks.
- Replacing the authenticated task-pane document destroys its JavaScript
  context and makes its memory-only MSAL cache unavailable without requiring
  Outlook to close. The application does not claim cryptographic process-memory
  zeroization or clearing of account or token state independently maintained by
  Outlook, NAA, Windows Web Account Manager, or another broker.
- The UI-less Check Status command runs in a manifest-declared short-lifetime
  runtime. Before signaling `event.completed()`, it releases OLHelper's
  reference to the MSAL instance so the memory cache is eligible for collection;
  Outlook controls final runtime destruction.
- Access-token lifetime is controlled by Microsoft Entra ID. OLHelper does not
  request a custom lifetime and does not retain an application timer-based copy.
- Memory-only caching reduces persistence but does not protect a token from
  malicious JavaScript executing in the active OLHelper origin.

## Application-enforced controls

### Mailbox operation boundaries

- The Graph client accepts only the exact endpoint shapes used for OLHelper
  folder, message-move, Inbox-scan, and Inbox-rule operations. It validates the
  HTTP method, exact query names and values, expected JSON body fields and
  values, and request options before obtaining a token.
- It rejects `/users`, `/me/drive`, Graph beta, alternate origins, credentials,
  fragments, unsupported headers or request options, broader `$select` values,
  unexpected filters, and unrelated mutation bodies.
- Validated absolute Graph v1 pagination links remain subject to the same
  method, path, field-selection, and pagination-parameter contract.
- Opaque folder and rule IDs are valid only inside approved path families.
  Semantic ownership and parent/target relationships are enforced by the
  folder and managed-rule modules; the central client cannot infer an opaque
  resource ID's Exchange parent from the URL alone.
- Folder, rule, message, and Tracking ID values are URL encoded.
- Inbox scans request only `id,subject`, use deterministic newest-first order,
  and stop after 250 messages.
- Tracking IDs require the canonical `TrackingID#` marker, 16–19 digits, and a
  numeric end boundary.

### User intent and stale-context protection

- Track, Archive, Delete, Reopen, and Repair require an in-pane confirmation describing
  the intended mailbox changes.
- The confirmation is bound to a snapshot of the Outlook item ID, converted
  Graph ID, and subject.
- Immediately before mutation, OLHelper revalidates that snapshot.
- The pinnable pane listens for `ItemChanged`; a change cancels confirmation,
  disables every action, reports why, and reloads the pane.
- Action controls start disabled and remain fail-closed if item-change
  protection cannot be registered.
- Task-pane initialization verifies that Office reports the Outlook host before
  enabling any action.

### Managed-rule ownership

OLHelper does not treat a display-name match as ownership. Before modifying or
deleting a candidate rule it verifies the complete managed fingerprint:

- Exact OLHelper-managed display name.
- Rule is writable and Graph reports no rule error.
- Exactly one expected `subjectContains` value.
- No additional configured conditions or exceptions.
- One non-empty move target.
- No forwarding, redirecting, deletion, marking, categorization, or other
  unexpected configured action.
- `stopProcessingRules` is not enabled.

A colliding or user-altered rule is rejected rather than adopted or modified.

### Web and deployment boundary

- CSP defaults all content to the OLHelper origin.
- Network access is limited to the OLHelper origin, Microsoft Graph, and
  Microsoft identity endpoints.
- Scripts are limited to the OLHelper origin and Microsoft-hosted Office.js.
- `object-src`, `base-uri`, and `form-action` are disabled.
- `frame-ancestors` permits only supported Outlook/Office hosts. OLHelper cannot
  use `X-Frame-Options: DENY` because Outlook must embed it.
- Referrers are disabled; HTTPS is enforced with HSTS; MIME sniffing is
  disabled; unnecessary browser capabilities are denied.
- HTML is served with `Cache-Control: no-store`.
- CI uses lockfile installation, production and full dependency audits,
  dependency review, CodeQL, and unified-manifest validation.
- GitHub Actions are pinned to immutable commit SHAs.
- The protected pilot workflow builds once, deploys that `dist` output, and
  publishes the matching unified package plus SHA-256 release hashes.
- After deployment, the workflow verifies the live HTML response against the
  reviewed CSP and security headers in `staticwebapp.config.json`; drift or
  incomplete propagation blocks package publication.

## Required tenant and operational controls

These controls are deployment requirements that application code cannot enforce 
directly, and are expected to be enforced at time of publication for the defined
security model:

1. Single-tenant Entra registration with only the two documented Graph
   delegated scopes.
2. Require assignment on the enterprise application and assign only approved
   pilot or production groups.
3. Grant admin consent only after identity, privacy, and security approval.
   Restrict ordinary user consent and use the admin-consent workflow.
4. Limit app-registration owners, enterprise-app administrators, Static Web App
   contributors, repository administrators, and deployment-environment
   approvers to named accountable owners.
5. Apply Conditional Access appropriate to the user population after report-only
   validation. 
6. Release branch protection: require current CI, code-owner review, secret
   scanning, and approval on the protected deployment environment.
7. Production host origin pinned in reviewed source; tenant ID and client ID
   kept in protected deployment configuration. Deployment token secret rotated
   on owner, exposure, or incident events.
8. Review assignments, delegated grants, app owners, repository owners, hosting
   roles, and production dependencies on a documented schedule.
9. Retain release commit, workflow run, package, and release hashes for each
   deployment. Promote the reviewed artifact rather than rebuilding from an
   unreviewed source state.

## Security release gates

Pending actions prior to official release:

1. Approval for `Mail.ReadWrite` and `MailboxSettings.ReadWrite`, using
   the permission-operation matrix above.
2. Enterprise-app assignment and admin-consent approval.
3. Confirmed production origin and NAA broker redirect URI.
4. Passes type checking, tests, production/full dependency audits, production
   build, and unified-manifest validation.
5. Automated confirmation that the production origin serves the reviewed CSP
   and security headers.
6. Exercise negative cases: malformed IDs, prefix overlap, renamed rules,
   changed selection after confirmation, partial sweep, Graph throttling, and
   rollback failure.
7. Record primary and backup owners, release approver, grant-review cadence,
   deployment-token rotation, rollback, and enterprise-app disable procedures.
8. Complete required privacy, accessibility, identity, and security reviews.

The repository defines code owners for release-sensitive workflow, manifest,
packaging, and security-policy files. The GitHub `main` ruleset must require
code-owner approval and current required checks, and the `pilot` environment
must require an independent reviewer with self-review prevention enabled.
These GitHub-hosted settings remain external controls and must be verified
before each production release.

## Residual risks requiring acceptance

- A delegated `Mail.ReadWrite` token can access more of the signed-in user's
  mailbox than OLHelper intentionally uses. No client-side check can narrows that
  OAuth authority. Existing mitigations heavily restrict the tokens local availabilty, but do not negate the risk associated with a stolen token.
- Compromise of the static origin, a runtime dependency, an approved release, or
  active WebView JavaScript could expose tokens and mailbox data within the
  delegated scopes. Existing mitigations are designed to minimize any supply-chain
  risks or vulnerabilities being introduced.
- Graph provides no transaction across folders, messages, and rules. OLHelper
  sequences operations, attempts limited rollback, reports partial success, and
  offers Repair, but rollback is not guaranteed. Existing validation checks are designed to minimize the probability of broken routing.
- The 250-message Inbox scan is intentionally bounded and may not move older
  matching mail. This introduces the possibility of missed email matches for the sake of avoiding full mailbox scans and reduced PII exposure potential.
- The delegated design has no server-enforced OLHelper-folder boundary. Exchange
  RBAC could provide a mailbox-population boundary, but would introduce a
  continuously privileged service identity and backend compromise surface.

## Emergency containment

1. Disable sign-in to the OLHelper enterprise application or remove all
   assignments.
2. Revoke delegated grants and affected user sessions/tokens according to the
   tenant incident procedure.
3. Remove the OLHelper app package from affected users.
4. Disable the deployment workflow and roll the Static Web App back to a known
   release or replace it with a non-operational holding page.
5. Rotate the Static Web Apps deployment token and review registration,
   repository, workflow, environment, and Azure role changes.
6. Preserve source, release hashes, workflow/Azure/Entra audit records, and
   relevant security events without collecting mailbox content.
7. Identify and disable or remove OLHelper-managed Inbox rules if deployed code
   may have been compromised.
8. Notify privacy, identity, security, and service owners under the applicable
   incident process before restoring service.
