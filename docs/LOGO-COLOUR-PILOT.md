# Logo colour decision — pilot completed

On 11 September 2026, the owner approved the outlined ARK and Artesian logos as the standard full-signature artwork for both light and dark backgrounds. Package 1.0.3 promotes those exact PNGs and retires the account-targeted test controls, alternative white images and switching rules.

## Standard artwork

The original vector letter shapes, dark fills and blue symbols are preserved. A pale `#F2F5F6` outline extends 0.4 CSS pixels beyond the dark lettering. One pixel of transparent padding on each side prevents clipping. The same image appears on both backgrounds; no theme detection is needed.

| Logo | PNG and matching SVG | Display size | PNG size |
|---|---|---|---|
| ARK | `public/assets/ark-logo.*` | 92 × 32 | 368 × 128 |
| Artesian | `public/assets/artesian-wordmark-color.*` | 64 × 20 | 256 × 80 |

The two PNGs total 23,867 bytes and are embedded as inline CID attachments. The PNGs exactly match the approved trial fallback assets. Matching SVGs are editable sources; email signatures use PNGs. Compact replies and forwards continue to use linked text without logo attachments.

## Why the switching experiment ended

The test supplied outlined dark-lettered and hidden white-lettered versions, with scoped dark-mode switching rules. Browser simulations displayed the intended pairs. In the actual Outlook Mac compose test, only the outlined versions appeared. The owner reported the same result in the received email in Apple Mail.

Inspection of that received email's raw MIME source found no style blocks, no media queries, no pilot classes, and only the two fallback images. Both white image elements and attachments were absent. The switching rules and artwork were lost along the tested delivery path; the received message had nothing to switch to. This evidence does not isolate the exact insertion, sending or transport step.

The owner preferred the outlined appearance in dark mode and the original-looking lettering on light backgrounds, so it became the standard design. Keep the raw email, screenshots and employee details outside this public repository. This acceptance covers the tested appearance, not every email client or scenario; the remaining checks are in [ACCEPTANCE.md](ACCEPTANCE.md).

## Publication and existing messages

After the reviewed change reaches `main`, the existing workflow publishes the updated signature bundle and panel. New compose sessions use the standard outlined artwork. Existing drafts can use **ARK signatures → Refresh this message**. If Outlook retains an old panel showing test controls, close and reopen Outlook. Sent messages retain their embedded images.

The manifest, Graph permissions and Microsoft 365 assignments are unchanged. No new Microsoft 365 deployment is required. To revert the artwork, restore the previous standard PNGs and display dimensions in a reviewed change; there is no per-user test preference to clear.
