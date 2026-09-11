# Outlook release acceptance and follow-up checks

Automated checks verify rendering logic and simulated Office API behaviour. They do not certify Outlook layout, authentication or installation.

On 9 September 2026, the owner accepted the transition from pilot to general release after confirming that a second assigned user received the add-in automatically. This is an operational release decision, not a claim that every client scenario below has passed. The 1.0.2 recovery patch passes 52 automated tests.

Record the client version, date, result and any screenshot for each scenario on **Mac, Windows new/classic as used by the team, and Outlook on the web**. On 8 September 2026, the pilot user approved two received external test messages: the revised full signature and the compact signature on a forwarded message. The screenshots confirm that typography, colours, logo sizing, phone spacing and the country-only footer now survive delivery. Keep screenshots and employee details outside this public repository. These results establish the v1.0.0 design baseline; they do not certify every Outlook client or every scenario below.

On 11 September 2026, the owner approved using the outlined trial logos as the standard for light and dark backgrounds. They were displayed in dark Outlook Mac composition and reported identically in the received Apple Mail test. The received MIME source retained exactly those two PNGs and removed the alternative white images and switching rules. Version 1.0.3 embeds the same approved PNG bytes directly, with no conditional switching. The 52 production regression tests remain; the 12 tests for the removed experiment were retired with that code.

| Scenario | Expected result | Result |
|---|---|---|
| Admin assignment | Saved deployment targets the approved user list | Last verified assignment: three named users, 9 September 2026. A wider employee list is being prepared for owner review before any assignment change; membership is recorded internally |
| Outlook availability | Add-in appears in assigned users’ Outlook Apps menu | First user confirmed on Mac, web and new Windows Outlook. Second user confirmed automatic arrival on 9 September 2026 after a deployment delay; third user’s availability remains unconfirmed |
| Outlook SSO and profile consent | Preview loads silently when consent/session permit; otherwise Continue with Microsoft 365 completes the required check; name/title/business phone match the directory | Profile-backed insertion confirmed by first pilot user; silent first use to be checked with the second user |
| New HTML message | ARK logo plus small linked Artesian wordmark after “Home of” | Passed: revised full signature approved in a received external self-test, 8 September 2026 |
| Reply / reply all / forward | Compact signature under the new reply, not at the bottom of quoted history | Forward passed: compact signature approved in a received external message, 8 September 2026. Reply and reply-all pending |
| New plain-text message | Readable contact details and co-brand line; no HTML or logo attachments | Pending |
| Internal and external delivery | Both small logo images display; clickable email/phone links work; layout survives transport | External visual checks passed for full and compact forward signatures after the formatting correction. Link-click verification and internal delivery pending |
| Sent Items | The compose-time signature is visible | Pending |
| Narrow window and dark mode | Text remains legible, no clipping; the small Artesian wordmark and reply text link are legible | Outlined full-signature logos accepted on 11 September 2026 following Outlook Mac dark composition and external Apple Mail receipt. Narrow-window, compact dark-mode and other client checks remain pending |
| Images blocked | Employee details remain readable; inline logo behaviour checked | Pending |
| Missing job title, phone or location | Missing lines are omitted | Pending |
| Repeated Refresh | Signature replaces itself; no duplicate logo attachments | Pending |
| Existing Outlook signature | Company signature replaces the signature in the new message without duplicates; saved personal templates remain intact | Pending |
| Switch approved sending identity | Current sender's identity; no borrowed title/phone from a different account | Pending |
| Unapproved sending domain | Notice appears and the user checks the remaining signature | Pending |
| Signed out / authentication expired | Recoverable rejection gets one silent retry; genuine sign-in checks require a user click; unrecovered failure preserves the existing signature | On 10 September a user reported recurrence after leaving Outlook open, recovering through the pane. Version 1.0.2 tests that failure/retry path with simulated provider responses. A real Mac idle-session retest is pending |
| Network unavailable / slow Graph | No message body overwrite, no send block, existing signature preserved | Pending |
| Central descriptor change | A newly composed message uses the published revision without employee editing | Pending |
| Central logo change | New messages get new embedded artwork; previous messages retain their original artwork | Pending |
| Pause and rollback | New compose events respect the pause; reverting a design restores its output after publication | Pending |

The renderer retains matching inline styles and scoped internal CSS. This accommodates the style-block loss reported for recent Outlook for Mac versions in [OfficeDev/office-js#6805](https://github.com/OfficeDev/office-js/issues/6805), while preserving the internal CSS path for other clients. Automated checks exercise removal of each styling form and mobile-number formatting. Browser checks confirmed matching typography with either styling form removed and no horizontal overflow at 272 pixels of content width. Real received-message checks now confirm the approved full and compact forward designs. Continue recording the remaining scenarios during the wider rollout. Second-user add-in availability is confirmed; it does not by itself confirm that user’s profile retrieval or sent-message formatting.
