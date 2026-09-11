# Account-targeted logo colour test

This experiment adds a manual test to the existing task pane for one configured pilot account. It does not change automatic signatures, the production signature bundle, saved preferences, the manifest, Microsoft 365 assignments or Graph permissions. The account's address is matched through a fixed SHA-256 digest; this is feature targeting, not an authentication boundary or anonymisation. Source and artwork are public, like the existing site. No employee directory or profile is published.

GitHub Pages has one production address and currently allows deployments only from `main`. A branch does not produce a mailbox-specific deployment. This branch therefore needs the normal reviewed merge to publish the account-targeted task-pane controls. The experimental implementation is separated into `src/experiments/` and its artwork into `public/experiments/logo-colours/`. No Pages policy or repository protection needs to change. No Microsoft 365 manifest upload is required.

## Test

1. After publication, reopen the ARK signatures panel; restart Outlook if it retains the older task-pane script.
2. In the assigned account, create a new HTML email using that account's own sending address. Wait for the normal automatic signature to appear.
3. Open ARK signatures, then choose **Insert test logos** under **Logo colour test**. The action changes only the signature slot of this draft and sends no message.
4. Send a test yourself to an internal Microsoft 365 mailbox and an external mailbox. Open the same received message in light and dark mode in Apple Mail and Outlook Mac, web and Windows.
5. Check both linked logos, hidden-image duplication, blocked images and the signature in quoted replies/forwards. Newly added reply signatures remain the current compact text version.
6. **Restore standard signature** restores the current company design in the draft. The normal **Refresh this message** button also restores it. New messages always start with the released design.

The main panel preview remains the standard company design; inspect the actual draft after clicking the trial button. If a slow automatic insertion finishes after the manual action, it may restore the standard signature: wait for the normal signature, then insert the test again. This pilot deliberately leaves the automatic runtime unchanged.

The trial and restore actions preserve existing attachments. Repeated insertion reuses its four named inline PNGs. Restoring the original HTML can leave unused trial PNG attachments in that draft; use a fresh draft for clean production mail.

## Rendering

The original vector glyph paths and blue symbols are preserved. Dark-lettered versions have a pale `#F2F5F6` outline extending 0.4 CSS pixels at display size. White-lettered variants are for supported dark-mode clients. Transparent padding prevents outline clipping. Display sizes including padding are ARK 92 × 32 and Artesian 64 × 20; PNGs are exported at 4×.

Four PNGs are embedded as inline CID attachments. Scoped `prefers-color-scheme:dark` rules and Outlook `data-ogsc`/`data-ogsb` selectors switch the image pair. White images default to `display:none` with zero inline and HTML dimensions; classic Outlook's MSO conditional comments exclude them. Inline fallbacks remain when a sending client strips style blocks. No script, sender-theme detection, tracking pixel, external image request or message-wide dark-mode rule is inserted into email.

Reference patterns: [Litmus dark-mode guide](https://www.litmus.com/blog/coding-emails-for-dark-mode), [Can I email compatibility tests](https://www.caniemail.com/features/css-at-media-prefers-color-scheme/). Support varies with client and version. A browser simulation cannot prove that Outlook preserves the CSS during delivery.

## Acceptance record

Automated checks cover account targeting, no additional work for non-pilot users, own-sender/profile checks, HTML-only new drafts, pause handling, duplicate clicks/attachments, standard restoration and safe default hiding. Browser checks and received-mail results are recorded separately. Do not treat publication or a successful preview as cross-client email acceptance.

On 11 September 2026, all 62 automated tests and Microsoft's manifest validation passed. Browser inspection of the actual rendered HTML confirmed exactly two visible logos in five cases: light and dark media-query contexts, removed style blocks, removed inline styles, and the Outlook `data-ogsc` selector. Media-query contexts were supplied through the containing iframe's `color-scheme`; no email client was simulated. Real sent/received Outlook testing is still pending.

The built automatic runtime, production signature bundle, manifest and Outlook well-known file are byte-for-byte identical to release 1.0.2. The four trial PNGs total 37,080 bytes; the rendered HTML is approximately 8 KB, below the existing 30 KB signature limit.

## Rollback

Remove the `installLogoPilot` import and call in `src/taskpane.js` and publish through the existing reviewed workflow. The files under `experiments/` can remain inert. No mailbox preferences or Microsoft 365 assignment changes need to be reversed.
