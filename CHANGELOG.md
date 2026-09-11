# Changelog

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
