# ARK signatures · Home of Artesian

One company signature, managed centrally for Mac, Windows and web Outlook, with mobile support ready for Android acceptance testing.

The shared design pairs ARK with “Home of” followed by a small official Artesian wordmark, linking to https://www.artesian.cloud/. The descriptor is **Energy markets. Managed data services. Technology.** New messages use the full design; replies and forwards use a compact version.

**Package v1.1.2.** Adds automatic mobile signatures and a connection/preview command accessible from received messages, with missing-marker and timeout recovery, ordered mobile insertion and safe failure references. Android device acceptance is pending. Enabling mobile requires updating the existing Microsoft 365 deployment to manifest **1.1.0.0**; publishing Pages alone does not enable it. The 1.1.1 and 1.1.2 patches require no further manifest update. The approved design, own-profile `User.Read` permission and verified SMTP alias handling remain. Microsoft 365 administrators control employee assignment separately; publishing a release does not change that list. Follow [Android setup and acceptance](docs/ANDROID-SETUP.md).

## How it works

1. GitHub holds the shared wording, links, HTML templates and logo PNGs.
2. GitHub Pages publishes a preview, the add-in and a versioned signature bundle.
3. When a new message, reply or forward is opened, the add-in checks the published bundle and reads the signed-in user's profile directly from Microsoft Graph.
4. It inserts the current signature and embeds the ARK logo and small Artesian wordmark in that email. Sent messages retain their original artwork.

Employee data and credentials are not stored in this repository. The add-in reads only the signed-in user's profile using delegated `User.Read`. The Outlook manifest requests `ReadWriteItem`, to insert a signature and an inline logo attachment in the current message. The application does not call APIs to send messages or read mail contents.

## Routine central changes

| Change | Where to edit |
|---|---|
| Descriptor, “Home of”, website links, phone/location preferences | `branding.json` |
| Layout | `templates/full.html` and `templates/reply.html` |
| Logos | `public/assets/ark-logo.png` and `public/assets/artesian-wordmark-color.png` |
| Employee name, job title, business phone | Microsoft 365 / Entra user profile |
| Shared country line (Ireland · Italy) | `locationLine` in `branding.json` |
| Pause or resume automatic insertion | `enabled` in `deployment.json` |

Save a reviewed change to `main`; the workflow runs tests, builds and publishes it. Future compose sessions fetch the new bundle after Pages finishes publishing. Existing drafts are not automatically rewritten; use **ARK signatures → Refresh this message**. A change does not rewrite already sent messages. Authentication, host caching and availability mean this is not an instantaneous push to every open Outlook window.

Matching editable SVGs accompany the PNGs. Display sizes, outline details and the completed email test are recorded in [the logo decision](docs/LOGO-COLOUR-PILOT.md). No theme detection or alternative hidden images are used.

## Deploy the release

Follow [Administrator setup](docs/ADMIN-SETUP.md), then complete [Outlook acceptance checks](docs/ACCEPTANCE.md).

Download versioned packages from [GitHub Releases](https://github.com/ARKlab/arklab.github.io/releases). [CHANGELOG.md](CHANGELOG.md) records the release baseline. GitHub Pages continues to publish approved changes from `main`; a release tag is a fixed reference for review and rollback.

The repository name `ARKlab/arklab.github.io` serves the add-in at the origin root, so Outlook can find `/.well-known/microsoft-officeaddins-allowed.json`. Using a project subfolder would require an additional root-site setup. There was no existing ARKlab Pages repository found when preparing this project.

## Local checks

```sh
npm ci
npm test
npm run build
npm run serve
```

The local preview is at `http://127.0.0.1:8766`. Outlook installation uses the published HTTPS site. `npm run validate:manifest` validates the built manifest; `npm run check:deployment` checks the published resources.

## Supported scope and limitations

- Desktop targets are current Microsoft 365 Outlook on Mac, Windows (new and classic) and the web. The manifest retains Mailbox 1.13 and the code checks NestedAppAuth 1.1. Older clients need updating.
- Mobile implementation targets Microsoft 365 Outlook on Android 4.2502.0 or later. The XML mobile declaration also enables iOS; real-device acceptance on both platforms is tracked separately. Mobile uses supported API exceptions without calling desktop-only body-format or attachment-list methods.
- Apple Mail as a sender, shared mailbox deployment and centrally enforced signatures at the mail server are outside this release.
- A sending address matching the signed-in user's primary address, sign-in address or directory SMTP alias receives that user's name, title and phone numbers, with the selected From address as its email link. An unverified address in an approved company domain receives only its Outlook display name and email. Other domains are skipped and a notice is shown. Verify sender switching before rollout.
- If authentication or the service is unavailable, the current signature is kept and a retry notice is shown. Maintain a local signature fallback if service availability is important.
- On mobile, open **ARK signatures** from a received message to check the connection, then create a new draft. There is no mobile compose-pane refresh button, and reopening a saved draft does not trigger automatic insertion. Quick replies must be expanded to see the signature.
- Users can still edit a message or remove a signature. This manages the default signature; it is not a compliance enforcement system.
- Missing fields are omitted. The country line is shared centrally; directory office locations are optional and off by default. Directory mobile numbers are not published unless an administrator explicitly enables them in `branding.json`.
- The official blue-and-dark Artesian wordmark appears at text scale, on a transparent background, and links to `https://www.artesian.cloud/`. Set `artesianWordmark:false` in `branding.json` for the live-text alternative. Replies always use linked text.

## References

- [Microsoft: Outlook signature example](https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/outlook-add-in-set-signature/)
- [Microsoft: NAA in Outlook events](https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/outlook-event-sso-naa/)
- [Microsoft: supported NAA clients](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets)
- [Microsoft: setSignatureAsync, including CSS constraints](https://learn.microsoft.com/en-us/javascript/api/outlook/office.body)
- [Microsoft: signed-in user profile and permissions](https://learn.microsoft.com/en-us/graph/api/user-get)
- [Microsoft: user properties, including SMTP proxy addresses](https://learn.microsoft.com/en-us/graph/api/resources/user)
- [Microsoft: centrally deploy add-ins](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins)
- [Microsoft: mobile event-based signatures](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/mobile-event-based)

The renderer retains matching inline formatting and scoped internal CSS so the signature remains styled when an Outlook client discards one form. The approved full and compact forward designs have been checked in received emails; the remaining client scenarios are tracked in the acceptance checklist.
