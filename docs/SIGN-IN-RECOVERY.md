# Sign-in recovery — 1.2.1

The observed failure was a Graph profile 401 after two token attempts in new Outlook for Windows, while the same colleague could use Outlook web. That narrows the investigation to the desktop authentication/request path, but does not prove token expiry, a cache defect or a Microsoft build regression. The old diagnostics cannot establish the specific cause.

## Recovery behavior

1. Automatic composition and opening the pane use silent authentication, retaining the existing total two-attempt budget. A first profile 401 requests a fresh token once.
2. A final profile 401 offers **Sign in again** in the pane. It does not open a popup automatically.
3. Clicking **Sign in again** uses Microsoft's interactive broker directly, then requests `/me` once. Microsoft decides which account, MFA or sign-in UI is needed; this does not guarantee a password prompt or clear Outlook's broker cache. Cancellation, a timeout or another rejection stops and shows support details. Another interaction requires another click.
4. Success enables **Refresh this message** on desktop. A preview never rewrites the draft by itself. On mobile, close the received-message pane and start a new message/reply/forward.

The recovery click reuses the settings already loaded in the pane to avoid an intervening network fetch before sign-in. Normal preview refresh and message insertion still fetch current settings. The same app, tenant and own-profile `User.Read` scope are used; this patch adds no permissions, storage of tokens or employee details, or sign-in redirects.

If Graph returns one unambiguous Bearer `insufficient_claims` challenge, its bounded, decoded `access_token` claims are carried to the next silent request and any subsequent user-triggered sign-in. Challenges are held in a private in-memory map and tied to the mailbox hint, tenant and app; they are not logged, serialized or persisted. Malformed, oversized, mixed/duplicate challenges are ignored. No authority or scope from the header is followed. This patch does not opt the client into additional capabilities such as `CP1`.

## Support details

The pane's **Support details** includes package version, UTC failure time, stage, token/profile attempt counts, silent/interactive mode, HTTP status, allowlisted Microsoft/Graph error codes and GUID request references when available. It also includes Office's reported platform and numeric host version. Unknown provider codes are omitted; raw response messages, tokens, employee information and claim values are excluded. Error bodies are limited to 16 KiB and the existing request timeout; claims are limited to 8 KiB decoded.

**OfficeOnline** can mean Outlook web or new Outlook for Windows. The Office host version is not necessarily the version in Outlook's About screen. Collect the app/client/WebView2 versions separately when reporting a problem. Missing HTTP headers or request IDs are reported by omission, not fabricated.

The background event and pane have separate runtime state. Opening the pane performs a new request; its support details describe that request, not the earlier compose failure. Background failures emit the same sanitized format to the add-in console. There is no remote telemetry or new logging service.

## Acceptance after publication

- Reopen ARK signatures after publication and verify **1.2.1** if support details appear. Restart Outlook if it still loads an older hosted bundle. No new manifest rollout is required for existing **1.2.0.0** installations.
- On the affected Windows account, reproduce the profile rejection, then choose **Sign in again**. Complete any Microsoft check. Confirm the preview loads; choose **Refresh this message** and verify the name/contact details belong to the intended sender.
- Verify a subsequent new draft and another after an idle period insert automatically. A single successful interactive recovery does not establish that the idle-session issue is permanently fixed.
- Cancel the interactive check once: the draft must stay unchanged and the pane must allow a fresh click. No popup should appear merely from creating another draft or opening the pane.
- If recovery still fails, share the new support details with IT, along with Outlook's About versions and whether web still works. Use the UTC timestamp and Microsoft correlation/request IDs to investigate; do not send tokens or raw authentication payloads.

Automated tests cover bounded retries, no background popups, cancellation, overlapping clicks, account changes, claims propagation, malformed/oversized input and diagnostic privacy, alongside the existing signature and mobile tests. They simulate the Microsoft broker; real-client acceptance remains pending.

## References

- [Microsoft: nested app authentication and interactive recovery](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in).
- [Microsoft: claims challenges](https://learn.microsoft.com/en-us/entra/identity-platform/claims-challenge).
- [Microsoft: Graph authentication errors](https://learn.microsoft.com/en-us/graph/resolve-auth-errors).
- [Microsoft: Office diagnostic context](https://learn.microsoft.com/en-us/javascript/api/office/office.contextinformation).
