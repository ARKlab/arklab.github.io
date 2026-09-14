# Outlook for Android setup and acceptance

Version 1.1.0 implements mobile support; version 1.1.1 fixes timeout and overlapping-insertion recovery, and version 1.1.2 handles an absent logo marker and adds safe failure references. Both patches use the same manifest. Automated checks and XML validation pass; real-phone acceptance is pending. The reported test client is Outlook for Android **5.2630.0 (72630118)**.

## What changes

New messages get the existing full signature with embedded outlined logos. Replies and forwards get the existing compact signature. From changes rerun insertion using the same verified-alias checks. Outlook SSO still reads only the signed-in user's profile with delegated `User.Read`.

The mobile path uses HTML directly because `body.getTypeAsync` is unsupported. It also avoids `getAttachmentsAsync`, which is not among the supported later mobile APIs. Successful logo attachments are recorded in per-item session data so another compose/From event can reuse them. Markers contain only asset filenames and an `added` flag; they contain no employee information or access tokens.

A new draft has no marker yet. Version 1.1.2 treats `KeyNotFound` (9050) from the marker read as this expected initial state. Other failures still stop insertion. If a refresh fails, the banner can include a safe reference such as `session-read/OUTLOOK_9051` or `signature-write/OUTLOOK_5001`; share that exact reference with IT. Opening the received-message pane diagnoses its own connection and preview, not a previous compose event.

Mobile insertions for the same item are queued within one runtime. Successful uploads are recorded even if the local deadline expired while the upload was pending; cancellation still prevents the next upload and signature write. Host termination or a failed marker write can still leave an unrecorded attachment. See [the review assessment](MOBILE-REVIEW.md).

## Publish and update Microsoft 365

1. Merge the reviewed change and wait for successful Pages publication. Verify the hosted code and `https://arklab.github.io/manifest.xml` are from the same build.
2. In Microsoft 365 **Settings → Integrated apps → ARK signatures**, update the existing app's manifest to the published **1.1.0.0** file. Retain the existing app ID and assigned users. Do not create a duplicate production add-in or change Graph permissions.
3. Verify the saved update and allow Microsoft to deliver it. Publishing GitHub Pages does not update the manifest stored by Microsoft 365. A manifest update can incur Microsoft's deployment propagation delay.
4. Restart Outlook on the test phone after delivery, then run the checks below. Updating the shared manifest enables mobile for all existing assigned users, not just the tester. Do not remove desktop users to narrow this test.

The old desktop installation continues to use the desktop path while manifest delivery is pending. The mobile declaration also covers iOS; it does not establish that iOS has passed acceptance.

The 1.1.1 and 1.1.2 patches change hosted code only; an existing 1.1.0.0 Microsoft 365 manifest does not need another update. Restart Outlook if it retains the previous runtime.

## Phone test

Use Outlook **4.2502.0 or later**, a Microsoft 365 business mailbox and a working connection. This minimum covers SSO, per-item session data and both supported compose/From events used by this app. No additional Microsoft permission or Entra app registration is required.

1. Open a received email and locate **ARK signatures** in the message's add-ins/menu. This opens a preview of **your own** profile. If Microsoft requests a sign-in check, choose **Continue with Microsoft 365**. The add-in does not modify the received email.
2. Close the panel and create a fresh email. Wait for the signature, send a test and inspect both the received email and Sent Items. Confirm names, job title, numbers, logo appearance and links.
3. Start a reply and a forward. Confirm the compact signature is beneath the new text. Expand a quick reply to full screen to see the signature.
4. Where the client exposes an alternate From address, select a verified company alias and confirm the displayed email changes while the owner's details remain. Switch back and check the result and attachments for duplicates.
5. Repeat after leaving Outlook idle, and after restarting it. Test light and dark mode and a narrow portrait display.
6. Check a temporary network failure. The add-in should preserve the current signature and complete its event without sending mail or overwriting the body. After reconnecting, start a new draft.

Keep employee screenshots and mail samples in internal records, outside this public repository.

## Mobile limitations

- The add-in runs automatically in compose; the manual pane is available when reading a message. There is no mobile **Refresh this message** button. After resolving a connection failure, create a new draft; reopening a saved draft does not fire the new-message event.
- A mobile event has a 60-second maximum and ends when the user sends or closes the message. Sending immediately can interrupt insertion; this app does not block sending or enforce signatures.
- Quick-reply signatures may be hidden until the composer is expanded.
- Offline profile retrieval cannot succeed. No employee directory cache or server-side signature service is introduced.
- The mobile API cannot enumerate attachments. Session markers prevent this add-in from repeatedly adding its own logo in the same session; they cannot detect a logo the user manually deletes. In that case, start a fresh draft.
- Outlook on iOS is enabled by the same manifest declaration but remains unverified. Apple Mail and other Android mail apps do not run this add-in.

## Acceptance record

| Check | Status |
|---|---|
| Generated mobile events, recovery command and desktop surfaces | Automated checks pass |
| Mobile-safe API path, duplicate prevention, partial failure and cancellation | Automated checks pass |
| Silent own-profile SSO, alias handling and explicit consent recovery | Simulated runtime checks pass |
| Microsoft XML manifest validation | Passed for 1.1.0.0 |
| Microsoft 365 manifest update and phone delivery | Pending |
| Real Android SSO and automatic compose/reply/forward insertion | Pending |
| Received-message appearance, clickable links and sender switching | Pending |
| Idle-session recovery and offline behaviour | Pending |
| iOS device acceptance | Pending |

## Microsoft references

- [Mobile event-based signatures and limitations](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/mobile-event-based)
- [Supported and unsupported mobile APIs](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-mobile-apis) — supported API exceptions can run even when the manifest requirement is higher than Mailbox 1.5; this manifest retains the desktop requirement of 1.13.
- [NAA client requirements](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets) — mobile SSO from 4.2433.0.
- [Mobile add-in commands](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/add-mobile-support)
- [Per-item session data](https://learn.microsoft.com/en-us/javascript/api/outlook/office.sessiondata)
