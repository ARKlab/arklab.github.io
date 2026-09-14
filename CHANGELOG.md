# Changelog

## 1.2.0 — Desktop shared-mailbox activation

The manifest omitted the shared-mailbox opt-in required to expose the add-in on Mac and some separately opened web/Windows setups. Enable `SupportsSharedFolders` on the desktop form factor and update the manifest to **1.2.0.0**.

- Retain the existing app ID, employee assignments, `ReadWriteItem` Office permission and own-profile Graph `User.Read` scope. No shared mailbox is added to the deployment assignment and no directory or mail permissions are introduced.
- Reuse the existing sender rules: an approved-domain shared From identity gets its Outlook display name and email, without the delegate's title or phone numbers. Personal accounts and verified aliases retain their own profile details.
- Cover the generated manifest, full/compact shared-sender insertion, compose preview and switching between shared and personal senders. Real client acceptance remains pending.

An administrator must update the existing Microsoft 365 deployment after Pages publication; publishing hosted code alone does not enable shared-mailbox activation. Microsoft's delivery delay applies. Outlook mobile does not support add-ins in shared mailboxes; the existing mobile declaration and personal-mailbox behavior are retained. See [shared-mailbox setup and acceptance](docs/SHARED-MAILBOXES.md).

## 1.1.2 — Missing mobile marker recovery and diagnostics

A new mobile draft could stop before its first logo upload when `sessionData.getAsync` returned Outlook's `KeyNotFound` code 9050. The marker is expected to be absent until insertion succeeds, but the previous flow treated that response as fatal.

- Treat only code 9050 from the mobile marker read as an absent marker. Continue to stop on other read errors and on all marker-write, attachment and signature-write failures.
- Update mobile test fixtures to return the missing-key error, including the compiled runtime and real generated production assets. Retain successful empty-return compatibility, repeated-insertion reuse, timeout handling and the per-item queue.
- Tag failing Office operations and expose allowlisted step names and numeric codes in mobile notifications and support details. Do not display provider messages, tokens, employee data or arbitrary symbolic error codes.

The regression reproduces `OUTLOOK_9050` against 1.1.1. This is a code defect consistent with the reported Android refresh message; the phone's actual failing code has not yet been collected, so successful device insertion still needs confirmation. Microsoft's current Android SDK maps 9050 to `KeyNotFound` ([public SDK](https://appsforoffice.microsoft.com/lib/1/hosted/outlook-android-16.00.debug.js)).

The manifest remains **1.1.0.0**. This is a Pages code update, with the same signature design, permissions and assignments, and requires no additional Microsoft 365 manifest update.

## 1.1.1 — Mobile insertion reliability

Fix mobile logo duplication after a local timeout and when compose/From-change handlers overlap in the same runtime.

- Record an attachment that already succeeded before checking cancellation, while preventing further uploads or signature writes after cancellation.
- Queue mobile insertions per item, so a pending older write cannot finish after a newer sender's write. A failed operation releases the queue; a queued request whose deadline expired stops before accessing the draft. Other drafts and the desktop path are independent.
- Test real generated logo filenames and attachment bytes, timeout recovery, concurrent updates, queued cancellation and independent drafts. Four regression scenarios fail against the previous merged code and pass with this patch.

See [the review assessment](docs/MOBILE-REVIEW.md) for the evidence on each Copilot finding, including why the claimed session-data key limit and icon deployment failure are not established by Microsoft's references. A hard host shutdown can still interrupt a marker write; the queue applies within one runtime. Real Android and iOS acceptance remains pending.

The manifest stays **1.1.0.0**, with the same signature artwork, permissions and assigned users. This patch requires Pages publication but no additional Microsoft 365 manifest update.

## 1.1.0 — Android acceptance candidate

Extend the add-in to Outlook mobile. Real Android deployment, SSO and sent-message checks remain pending; iOS is enabled by the same XML mobile declaration but is also unverified on a device.

- Add `MobileFormFactor` with new-message and From-change events, and a received-message **ARK signatures** command for connection checks and preview.
- Use HTML insertion on mobile without calling the unsupported body-format and attachment-list APIs. Track successfully attached logos in per-item session data to avoid duplicates across repeated events; no employee details or tokens are stored there.
- Preserve full/compact signature designs, owner-only Graph SSO, verified aliases and sender-change cancellation. Keep desktop behaviour and the same app identity and permissions.
- Hide received-message insertion in the mobile pane and direct recovery to a new draft after reconnection.

Validation: 68 automated checks pass, including mobile runtime, alias switching, replies/forwards, partial attachment recovery, cancellation, draft isolation, consent and generated manifest coverage. Microsoft's manifest validator accepts the XML. These checks simulate Office APIs; they do not certify a real phone's behaviour.

**Requires a Microsoft 365 manifest update to 1.1.0.0.** Pages publication updates the hosted code only. Keep the existing app and employee assignment; update its manifest after publication. Microsoft deployment propagation may apply. See [Android setup and acceptance](docs/ANDROID-SETUP.md).

## 1.0.4 — 14 September 2026

Fix signature insertion when sending from an `ark-energy.it` alias. Previously that domain was rejected, and alternate addresses were not matched to the signed-in user's full profile.

- Approve `ark-energy.it` alongside the existing company domains.
- Read SMTP aliases from `proxyAddresses` on the existing Microsoft Graph `/me` request. A matching alias retains the owner's name, title and phone numbers while displaying and linking the selected From address.
- Keep sign-in tied to the mailbox owner. Unverified addresses retain the name/email-only fallback; non-SMTP contact addresses cannot establish ownership. The final From-address check still cancels insertion if the sender changes during profile retrieval.

Validation: 56 automated tests pass, including verified aliases in full, compact and plain-text signatures, case handling, rejected ownership matches, sender changes and the compiled automatic From-change flow. After publication on 14 September, the owner confirmed that sending from the reported alias works.

The manifest remains `1.0.0.0`, with delegated `User.Read` and no directory-wide access. No Microsoft 365 redeployment, new permissions or employee-assignment change is required. Outlook may need to be restarted if it retains the previous hosted code.

## 1.0.3 — 11 September 2026

Standardise the approved outlined ARK and Artesian logos in all full signatures. The same artwork keeps the original dark lettering on light backgrounds and reveals a fine pale outline on dark backgrounds.

- Promote the exact PNGs approved during the logo trial, with matching editable SVGs. Preserve their transparent padding at ARK 92 × 32 and Artesian 64 × 20 CSS pixels.
- Embed only the two standard logos. Retire the account-targeted test controls, hidden white images and conditional switching rules. The trial's received email lost those rules and white images along the Outlook-to-iCloud delivery path.
- Keep the compact reply/forward design, company wording, profile permissions and automatic sign-in recovery.

Validation: 52 regression tests pass. The owner approved the outlined appearance after Outlook Mac composition and an external received-message test in Apple Mail. The standard PNG bytes match those accepted trial assets. Broader client coverage remains in the acceptance checklist.

The manifest remains `1.0.0.0`; no Microsoft 365 redeployment or assignment change is required. After Pages publication, new compose sessions fetch the updated bundle. Existing drafts can use **Refresh this message**. Restart Outlook if a cached pane still shows the retired test controls.

## 1.0.2 — 10 September 2026

Automatic insertion now attempts silent recovery when Outlook rejects a sign-in that can succeed on retry. Previously, these token-acquisition exceptions stopped insertion immediately, even though opening the pane and refreshing could work.

- Retry a completed broker rejection (7000024) or an identified transient sign-in error once after a one-second delay, requesting a fresh token. Token-acquisition recovery and Graph 401 recovery share a two-attempt limit.
- Keep required interaction behind an explicit click. Do not retry unknown errors, policy/consent failures, an unsuccessful popup, or a native authentication request that is still pending.
- Extend the background event limit from 18 to 60 seconds, and limit authentication and Graph requests to the remaining budget with time reserved for profile retrieval and signature insertion. Normal successful requests incur no delay.
- Recognize numeric Microsoft error codes and the known broker-rejection description. Preserve allowlisted sign-in/suberror codes, token-attempt count and a request correlation reference when available. Background warnings log only safe support references; pane errors show those references in Support details. No tokens or employee profiles are logged.

Validation: 52 automated tests pass, including complete automatic insertion after a simulated initial broker failure, repeated rejection, cancellation, shared retry limits and diagnostic privacy. The Microsoft-side defect is not reproduced by these tests; acceptance after Outlook has been left open remains pending.

The manifest remains `1.0.0.0`. This hosted-code update requires no new Microsoft 365 assignment or permissions. The approved full/compact signatures are unchanged.

## 1.0.1 — 9 September 2026

**General release of the 1.0 series.** The owner accepted the end of the pilot on 9 September 2026 after confirming automatic delivery to a second assigned user. Wider employee assignment is managed separately in Microsoft 365 and awaits review of the proposed employee list.

Includes the approved v1.0.0 signature design and improved recovery and diagnosis when an installed add-in cannot refresh its Microsoft 365 profile.

- Retry a Graph 401 once with a fresh access token; an explicit Retry connection after a sign-in failure also bypasses the token cache.
- Bound silent sign-in to 15 seconds and initialization to 10 seconds. User-requested Microsoft interaction has a separate two-minute deadline. A still-pending native sign-in cannot start duplicate requests; late results cannot load a profile or insert a signature after timeout.
- Cancel timed-out HTTP requests, retaining the existing eight-second HTTP and eighteen-second background insertion deadlines.
- Distinguish Microsoft broker rejection (AADSTS7000024), sign-in timeout, settings download, profile retrieval and Outlook insertion failures. Optional support details expose only the application version, time, step, safe result code and Microsoft correlation reference.
- Keep Retry available after an initial settings failure, disable insertion after a failed refresh, and clear the earlier compose warning after successful manual insertion.

Validation: 41 automated authentication, timeout, UI recovery, privacy, sender-isolation and signature-rendering tests pass. These simulate failure/recovery paths; they do not establish a permanent fix for native Outlook broker error 7000024. If a fresh-token retry still fails, fully quitting and reopening Outlook remains the recovery step. Microsoft interaction is still requested only after an explicit click and an interaction-required error, following [Microsoft's NAA example](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in).

The hosted application changes; the manifest remains `1.0.0.0`, with the same permissions and pilot assignments. No Microsoft 365 redeployment is required. The second user’s catalogue entry subsequently appeared without another assignment change; this patch does not change Microsoft 365 catalogue delivery.

## 1.0.0 — 8 September 2026

First official release of ARK signatures, home of Artesian. Rollout remains limited to designated Microsoft 365 pilot users.

- Approved full signature for new messages: ARK logo, linked official Artesian wordmark, “Energy markets. Managed data services. Technology.” and “Ireland · Italy”. Both logos link to their websites, without a duplicate website footer.
- Compact reply/forward signature with bold blue ARK and linked Artesian text.
- Shared templates and embedded PNG artwork published from GitHub Pages; employee profile data is read directly from Microsoft 365 and is not stored on GitHub.
- Outlook SSO through nested app authentication, delegated `User.Read`, and explicit interaction only when Microsoft requires consent or a sign-in check.
- Automatic insertion for assigned users, centrally controlled pause/resume, and manual refresh of an existing draft.
- Matching inline and scoped internal CSS for Outlook compatibility; readable Irish and Italian mobile-number spacing with dialable phone links.
- Sender-identity checks, escaped profile fields, plain-text output, signature-slot updates and reuse of existing managed logo attachments.

Validation: 23 automated tests pass. The pilot user approved external received-email visuals for the full signature and compact forwarded signature. Remaining scenarios and client coverage are recorded in [the acceptance checklist](docs/ACCEPTANCE.md). Signatures remain editable by employees; server-side enforcement and mobile Outlook are outside this release's scope.

The Outlook manifest remains version `1.0.0.0`. This release records the approved package and documentation; adding a pilot user uses the existing Microsoft 365 assignment and does not require a new manifest.
