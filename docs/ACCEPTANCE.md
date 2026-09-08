# Outlook pilot acceptance

Automated checks verify rendering logic and simulated Office API behaviour. They do not certify Outlook layout, authentication or installation.

Record the client version, date, result and any screenshot for each scenario on **Mac, Windows new/classic as used by the team, and Outlook on the web**. No real Outlook acceptance result has been recorded yet.

| Scenario | Expected result | Result |
|---|---|---|
| Admin assignment | Saved deployment targets the agreed pilot user only | Passed in Microsoft 365 admin centre, 8 September 2026 |
| Outlook availability | Add-in appears in the pilot user's Outlook Apps menu | Pending propagation; absent on initial web check |
| Outlook SSO and profile consent | Preview loads silently when consent/session permit; otherwise Continue with Microsoft 365 completes the required check; name/title/business phone match the directory | Pending |
| New HTML message | ARK logo plus small linked Artesian wordmark after “Home of” | Pending |
| Reply / reply all / forward | Compact signature under the new reply, not at the bottom of quoted history | Pending |
| New plain-text message | Readable contact details and co-brand line; no HTML or logo attachments | Pending |
| Internal and external delivery | Both small logo images display; clickable email/phone links work; layout survives transport | Pending |
| Sent Items | The compose-time signature is visible | Pending |
| Narrow window and dark mode | Text remains legible, no clipping; the small Artesian wordmark and reply text link are legible | Pending |
| Images blocked | Employee details remain readable; inline logo behaviour checked | Pending |
| Missing job title, phone or location | Missing lines are omitted | Pending |
| Repeated Refresh | Signature replaces itself; no duplicate logo attachments | Pending |
| Existing Outlook signature | Company signature replaces the signature in the new message without duplicates; saved personal templates remain intact | Pending |
| Switch approved sending identity | Current sender's identity; no borrowed title/phone from a different account | Pending |
| Unapproved sending domain | Notice appears and the user checks the remaining signature | Pending |
| Signed out / authentication expired | Existing signature remains; connection/retry guidance appears | Pending |
| Network unavailable / slow Graph | No message body overwrite, no send block, existing signature preserved | Pending |
| Central descriptor change | A newly composed message uses the published revision without employee editing | Pending |
| Central logo change | New messages get new embedded artwork; previous messages retain their original artwork | Pending |
| Pause and rollback | New compose events respect the pause; reverting a design restores its output after publication | Pending |

If a client cannot preserve the scoped internal CSS through `setSignatureAsync`, hold rollout and adjust the renderer based on that client’s observed behaviour. Do not treat the browser preview as an email-client test.
