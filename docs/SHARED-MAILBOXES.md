# Shared-mailbox setup and acceptance

Version **1.2.0**, manifest **1.2.0.0**, enables `SupportsSharedFolders` inside `DesktopFormFactor` in the XML 1.1 version overrides. The earlier manifest omitted this opt-in, so Outlook could hide the add-in when a shared mailbox was opened separately. This is an activation change; the existing renderer and signature insertion code are retained.

## Client coverage

| Outlook client | Expected coverage after manifest delivery |
|---|---|
| Mac, shared mailbox opened through delegated access | Desktop opt-in is required; verify new compose, reply and the ARK signatures command |
| Classic Windows, shared mailbox added through delegated access | Uses the add-in assigned to the employee's primary mailbox; verify compose and sender changes |
| New Windows, shared mailbox promoted or added as an account | Desktop opt-in is required for this shared context; verify with the employee's delegated access |
| Web, Open another mailbox in a separate tab/window | Desktop opt-in is required; verify compose and the command in that window |
| Android / iOS, shared mailbox | Not supported by Outlook; this release cannot enable automatic signatures there |

Use supported Microsoft 365 clients with the app's existing Mailbox 1.13 and NestedAppAuth 1.1 capabilities. See [Microsoft's shared-mailbox support and setup guide](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/delegate-access) and [From-change behavior](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/onmessagefromchanged-onappointmentfromchanged-events). These document platform support; they do not certify this app's real-client results.

Microsoft's [add-in availability table](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/delegate-access#add-ins-in-shared-mailbox-scenarios), checked on 2026-09-14, explicitly covers **Open another mailbox** in a separate web tab/window and promoted shared mailboxes in new Windows. It requires shared-mailbox manifest support and installation in the employee's primary mailbox for compose in those contexts. These remain acceptance targets, not confirmed device results.

## Signature identity

The selected **From** address controls the signature. A sender on an approved company domain that does not match the signed-in employee's primary address or directory SMTP aliases receives its Outlook display name and email, plus the common ARK/Artesian branding. If Outlook supplies no display name, the email is used. The employee's personal job title and phone numbers are omitted. New messages use the full layout; replies and forwards use the compact layout.

SSO still obtains the signed-in employee's own `/me` profile with `User.Read`. The add-in does not retrieve the shared mailbox's directory profile, phone numbers or messages through Graph. Its Office permission remains `ReadWriteItem`, and it writes the signature slot of the current compose item. Returning to the employee's personal From address restores their own profile details. Unapproved domains remain rejected.

The sign-in hint remains the user identity Outlook provides; it is not changed to the selected shared From address. Real shared-context SSO must be checked on each client. If authentication fails, record the support details rather than signing in directly with shared-mailbox credentials.

## Update the existing deployment

1. Merge the reviewed PR, verify successful Pages publication and check that the hosted manifest is **1.2.0.0**.
2. In Microsoft 365 **Settings → Integrated apps → ARK signatures**, update the existing app's manifest from `https://arklab.github.io/manifest.xml` (or upload that downloaded file). Keep the existing app ID and employee assignment list.
3. Keep the add-in assigned to employees' primary mailboxes. Microsoft advises against assigning it directly to shared mailboxes. Existing Exchange delegation and Send As/Send on Behalf permissions remain separate and are not altered by this update.
4. Verify the saved update, allow Microsoft's delivery delay and restart Outlook before creating a new draft from the separately opened shared mailbox. Pages publication alone cannot update the manifest stored in Microsoft 365.

## Test before confirming coverage

On Mac and the Windows variant used by the employee, open a shared mailbox separately using delegated access. Also test web **Open another mailbox** if used by the team. Keep mailbox addresses, screenshots and assignment evidence in internal records.

| Check | Status |
|---|---|
| Desktop opt-in in generated XML, retained desktop events/commands and mobile declaration | Automated checks pass |
| Microsoft XML manifest validation | Passed for 1.2.0.0 |
| Full/compact shared-sender output, compose preview and switching back to the employee | Simulated checks pass; no borrowed title or phone |
| Microsoft 365 manifest 1.2.0.0 delivered | Pending |
| ARK signatures command visible in the separately opened shared mailbox | Pending |
| Silent SSO and automatic signature in a fresh shared-mailbox draft | Pending |
| Reply/forward, repeated manual refresh, switch between personal/shared From | Pending |
| Received message retains mailbox identity, links and branding | Pending |
| Existing personal-mailbox desktop and Android signatures still work | Real-client regression pending |

If the command remains absent after delivery, verify the actual client, mailbox-opening method and employee assignment. Microsoft also documents a limitation for mailboxes hidden from address lists: the current item may be unavailable to the add-in. Diagnose that separately; do not automatically change mailbox visibility or grant broader permissions.

For shared-mailbox messages on mobile, use a manually entered signature until a separately approved mobile-compatible approach is adopted. This release does not introduce a server-side signature service.
