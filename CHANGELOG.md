# Changelog

## 1.0.1 — 9 September 2026

Improve recovery and diagnosis when an installed add-in cannot refresh its Microsoft 365 profile.

- Retry a Graph 401 once with a fresh access token; an explicit Retry connection after a sign-in failure also bypasses the token cache.
- Bound silent sign-in to 15 seconds and initialization to 10 seconds. User-requested Microsoft interaction has a separate two-minute deadline. A still-pending native sign-in cannot start duplicate requests; late results cannot load a profile or insert a signature after timeout.
- Cancel timed-out HTTP requests, retaining the existing eight-second HTTP and eighteen-second background insertion deadlines.
- Distinguish Microsoft broker rejection (AADSTS7000024), sign-in timeout, settings download, profile retrieval and Outlook insertion failures. Optional support details expose only the application version, time, step, safe result code and Microsoft correlation reference.
- Keep Retry available after an initial settings failure, disable insertion after a failed refresh, and clear the earlier compose warning after successful manual insertion.

Validation: automated authentication, timeout, UI recovery, privacy, sender-isolation and signature-rendering tests pass. These simulate failure/recovery paths; they do not establish a permanent fix for native Outlook broker error 7000024. If a fresh-token retry still fails, fully quitting and reopening Outlook remains the recovery step. Microsoft interaction is still requested only after an explicit click and an interaction-required error, following [Microsoft's NAA example](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in).

The hosted application changes; the manifest remains `1.0.0.0`, with the same permissions and pilot assignments. No Microsoft 365 redeployment is required. The separate missing-add-in catalogue issue is not addressed by this patch.

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
