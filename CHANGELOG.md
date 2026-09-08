# Changelog

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
