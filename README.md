# ARK signatures · Home of Artesian

One company signature, managed centrally for Mac, Windows and web Outlook.

The shared design pairs ARK with “Home of” followed by a small official Artesian wordmark, linking to https://www.artesian.cloud/. The descriptor is **Energy markets. Managed data services. Technology.** New messages use the full design; replies and forwards use a compact version.

**1.0 general release — v1.0.2, 10 September 2026.** This patch adds one silent fresh-token retry for recoverable sign-in rejections during automatic insertion, with bounded waits and safe diagnostics. The approved full and compact signatures remain unchanged. The app uses Outlook SSO and delegated `User.Read` for the signed-in user's own profile. Microsoft 365 administrators control employee assignment separately; publishing a release does not change that list. Remaining client scenarios are recorded in the acceptance checklist.

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

- Supported targets are current Microsoft 365 Outlook on Mac, Windows (new and classic) and the web. The manifest uses Mailbox 1.13 and the code checks NestedAppAuth 1.1. Older clients need updating.
- Mobile, Apple Mail, shared mailbox deployment and centrally enforced signatures at the mail server are outside this release.
- An alternate sending address in an approved company domain receives its Outlook display name and email, with no borrowed job title or phone number. Other domains are skipped and a notice is shown. Verify sender switching before rollout.
- If authentication or the service is unavailable, the current signature is kept and a retry notice is shown. Maintain a local signature fallback if service availability is important.
- Users can still edit a message or remove a signature. This manages the default signature; it is not a compliance enforcement system.
- Missing fields are omitted. The country line is shared centrally; directory office locations are optional and off by default. Directory mobile numbers are not published unless an administrator explicitly enables them in `branding.json`.
- The official blue-and-dark Artesian wordmark appears at text scale, on a transparent background, and links to `https://www.artesian.cloud/`. Set `artesianWordmark:false` in `branding.json` for the live-text alternative. Replies always use linked text.

## References

- [Microsoft: Outlook signature example](https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/outlook-add-in-set-signature/)
- [Microsoft: NAA in Outlook events](https://learn.microsoft.com/en-us/samples/officedev/office-add-in-samples/outlook-event-sso-naa/)
- [Microsoft: supported NAA clients](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/nested-app-auth-requirement-sets)
- [Microsoft: setSignatureAsync, including CSS constraints](https://learn.microsoft.com/en-us/javascript/api/outlook/office.body)
- [Microsoft: signed-in user profile and permissions](https://learn.microsoft.com/en-us/graph/api/user-get)
- [Microsoft: centrally deploy add-ins](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins)

The renderer retains matching inline formatting and scoped internal CSS so the signature remains styled when an Outlook client discards one form. The approved full and compact forward designs have been checked in received emails; the remaining client scenarios are tracked in the acceptance checklist.
