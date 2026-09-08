# Outlook pilot acceptance

Automated checks verify rendering logic and simulated Office API behaviour. They do not certify Outlook layout, authentication or installation.

Record the client version, date, result and any screenshot for each scenario on **Mac, Windows new/classic as used by the team, and Outlook on the web**. On 8 September 2026, the pilot user confirmed successful insertion and an external self-test delivery. The received screenshot showed that typography, colours and spacing had been lost; visual acceptance therefore requires another received test after the formatting correction.

| Scenario | Expected result | Result |
|---|---|---|
| Admin assignment | Saved deployment targets the agreed pilot user only | Passed in Microsoft 365 admin centre, 8 September 2026 |
| Outlook availability | Add-in appears in the pilot user's Outlook Apps menu | Confirmed by pilot user, 8 September 2026 |
| Outlook SSO and profile consent | Preview loads silently when consent/session permit; otherwise Continue with Microsoft 365 completes the required check; name/title/business phone match the directory | Pending |
| New HTML message | ARK logo plus small linked Artesian wordmark after “Home of” | Insertion confirmed by pilot user; revised styling awaiting received-message retest |
| Reply / reply all / forward | Compact signature under the new reply, not at the bottom of quoted history | Pending |
| New plain-text message | Readable contact details and co-brand line; no HTML or logo attachments | Pending |
| Internal and external delivery | Both small logo images display; clickable email/phone links work; layout survives transport | External self-test received with both logos; initial styling failed. Corrected rendering and link actions need retest; internal delivery pending |
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

The renderer now retains matching inline styles and scoped internal CSS. This accommodates the style-block loss reported for recent Outlook for Mac versions in [OfficeDev/office-js#6805](https://github.com/OfficeDev/office-js/issues/6805), while preserving the internal CSS path for other clients. Automated checks exercise removal of each styling form and mobile-number formatting. Browser checks confirmed matching typography with either styling form removed and no horizontal overflow at 272 pixels of content width. These are simulations, not Outlook client certification; retain the pilot assignment until the received-message retest passes.
