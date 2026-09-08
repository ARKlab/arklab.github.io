# Administrator setup

This is a reviewable pilot package. Publishing GitHub Pages does not register or install the Microsoft 365 add-in.

## Hosting prerequisite

ARKlab must allow public Pages publication before this repository can serve the add-in. An organisation owner controls this under **Organisation settings → Member privileges → Pages creation → Public**. This is an organisation-wide policy, not a setting limited to the signature repository; obtain the owner's approval before changing it. See [GitHub's instructions](https://docs.github.com/en/organizations/managing-organization-settings/managing-the-publication-of-github-pages-sites-for-your-organization).

Once allowed, open **Repository settings → Pages**, select **GitHub Actions** as the source, then run **Check and publish signatures** from the Actions tab. The workflow includes hidden files so Outlook's required `/.well-known` file is deployed. The build can pass while publication is still blocked by organisation policy.

After initial publication, the organisation's restriction on creating Pages sites can be restored. GitHub states that existing published sites remain published; routine signature changes use this existing site.

## 1. Register the application in your Microsoft 365 tenant

In Microsoft Entra admin center, open **App registrations → New registration**.

| Setting | Value |
|---|---|
| Name | ARK central signatures |
| Account type | Accounts in this organisational directory only (single tenant) |
| Redirect platform | Single-page application (SPA) |
| Redirect URI | `brk-multihub://arklab.github.io` |
| Microsoft Graph permission | Delegated `User.Read` |
| Client secret | None — this is a public client using nested app authentication |

Record the **Directory (tenant) ID** and **Application (client) ID**. They are identifiers, not secrets. Set `tenantId` and `clientId` in `deployment.json`. Do not put tokens, passwords, client secrets or employee records into this project.

Have an authorised administrator review and grant the required consent under your tenant policy. `User.Read` permits reading the signed-in user's profile. This application does not require `User.Read.All`, `Directory.Read.All`, Graph mail access or a backend with application permissions.

See [Microsoft's NAA registration guide](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in). The broker redirect uses the origin only, with no trailing path. If the hosting domain changes, update the app registration and `siteUrl`, then rebuild and redeploy the manifest.

## 2. Confirm brand and directory settings

- The confirmed `artesianWebsite` is `https://www.artesian.cloud/`; update it in `branding.json` if the destination changes.
- Confirm the sending domains in `approvedSenderDomains`.
- Check the pilot employees' display name, job title, `businessPhones`, primary `mail` and `officeLocation` in Microsoft 365. Microsoft 365 administration is the central place to change those details.
- `mobilePhone` is enabled for the single-user pilot, whose contact numbers have been approved. Before expanding the assignment, confirm that this field holds an approved business contact number for every affected employee, or set `includeMobilePhone:false`. Numbers appear without office/mobile labels.
- The shared location line is **Ireland · Italy**. Change `locationLine` centrally if required. Directory office locations are off by default; clear `locationLine` and enable `showOfficeLocation` to use them. Review any approved company/legal footer requirements before general rollout.

## 3. Publish and verify

Commit the identifiers and reviewed settings. Keep `enabled` false while setting up. The GitHub workflow tests and publishes to `https://arklab.github.io/`.

Run `npm run check:deployment`. Confirm these URLs work without sign-in:

- `https://arklab.github.io/manifest.xml`
- `https://arklab.github.io/taskpane.html`
- `https://arklab.github.io/.well-known/microsoft-officeaddins-allowed.json`

The well-known file must be at the origin root and must allow `https://arklab.github.io/runtime.js`. This is required by Outlook for event-based authentication/CORS. See [Microsoft's requirement](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/use-sso-in-event-based-activation).

The published site contains application code, generic templates, branding and app identifiers. Employee profile data travels directly between the signed-in Outlook client and Microsoft Graph. The application doesn't send that data to GitHub.

## 4. Install for a small pilot group

In **Microsoft 365 admin center → Settings → Integrated apps**, upload the custom Outlook add-in using the published `manifest.xml`, review its requested permissions, and assign it only to a small pilot group initially.

The manifest uses the XML add-in format to support Mac as well as Windows and web Outlook. Its Office permission is **ReadWriteItem**: the ability to update the current message and attach the two small logo images. The code changes only the signature slot; it never sends messages. An administrator must review the real consent screens before installation.

Use a pilot mailbox on each of Mac, Windows and web Outlook. In a message, open **ARK signatures**, choose **Connect Microsoft 365**, and complete any required sign-in/consent. The preview should contain the correct profile. Event code uses silent authentication; it cannot display sign-in popups itself.

When ready to test automatic insertion, set `enabled` to `true` and publish. The build refuses to enable the pilot without configured tenant and client IDs. The add-in only runs for users to whom it has been assigned.

## 5. Roll out after acceptance

Complete [ACCEPTANCE.md](ACCEPTANCE.md) with real send/receive tests. After the pilot is accepted, expand the deployment group. Keep a documented local signature fallback until availability and client coverage are established. The Microsoft 365 deployment confirmation advises allowing up to 72 hours for the add-in to appear and says users may need to relaunch Microsoft 365.

For routine wording or branding updates, edit `branding.json` or the templates and publish. Administrator reinstall is normally unnecessary for content changes. Manifest changes — new URLs, IDs, permissions, events or requirements — need a manifest version bump and an administrator-managed update.

## Pause and rollback

- To pause automatic insertion, set `enabled:false` and publish. After the new bundle becomes available, new compose events retain the existing client signature.
- To roll back a design, revert its GitHub commit and republish.
- To remove the pilot, unassign or remove the add-in in Integrated apps and restore the usual Outlook signature defaults.
- An already open draft may contain a previously inserted signature; pausing doesn't edit existing drafts or sent emails.

## Current pilot status

The single-tenant **ARK central signatures** app is registered. Its identifiers are configured in `deployment.json`, and its only requested Graph permission is delegated `User.Read`. No client secret is used.

On 8 September 2026, Microsoft 365 reported **Deployment completed**. The saved assignment was verified as **Just me**, matching the agreed pilot account, with no other users or groups selected. The user explicitly approved the installation capabilities: `ReadWriteItem`, `SendReceiveData` and `ProfileAccess`.

The initial Outlook web Apps menu has not yet shown the new add-in. User connection, preview and real Outlook acceptance remain pending. Automatic insertion remains paused. Employee records and the pilot user's contact details are not stored in this repository.
