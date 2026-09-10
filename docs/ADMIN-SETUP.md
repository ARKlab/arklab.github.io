# Administrator setup

This is the 1.0 general release, package v1.0.2. The owner accepted the end of the pilot on 9 September 2026. Publishing GitHub Pages or creating a GitHub release does not register, install or assign the Microsoft 365 add-in.

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
- Check the assigned employees' display name, job title, `businessPhones`, primary `mail` and `officeLocation` in Microsoft 365. Microsoft 365 administration is the central place to change those details.
- `mobilePhone` is enabled in the current configuration. Before expanding the assignment, confirm that this field holds an approved business contact number for every affected employee, or is blank; otherwise set `includeMobilePhone:false`. Numbers appear without office/mobile labels. The directory profile's Business phone field is sufficient when Mobile phone is blank.
- The shared location line is **Ireland · Italy**. Change `locationLine` centrally if required. Directory office locations are off by default; clear `locationLine` and enable `showOfficeLocation` to use them. Review any approved company/legal footer requirements before general rollout.

## 3. Publish and verify

Commit the identifiers and reviewed settings. Keep `enabled` false while setting up. The GitHub workflow tests and publishes to `https://arklab.github.io/`.

Run `npm run check:deployment`. Confirm these URLs work without sign-in:

- `https://arklab.github.io/manifest.xml`
- `https://arklab.github.io/taskpane.html`
- `https://arklab.github.io/.well-known/microsoft-officeaddins-allowed.json`

The well-known file must be at the origin root and must allow `https://arklab.github.io/runtime.js`. This is required by Outlook for event-based authentication/CORS. See [Microsoft's requirement](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/use-sso-in-event-based-activation).

The published site contains application code, generic templates, branding and app identifiers. Employee profile data travels directly between the signed-in Outlook client and Microsoft Graph. The application doesn't send that data to GitHub.

## 4. Install or use the existing deployment

In **Microsoft 365 admin center → Settings → Integrated apps**, upload the custom Outlook add-in using the published `manifest.xml`, review its requested permissions, and assign it to the approved employee list. For the existing ARK deployment, update its user assignment rather than uploading another copy.

The manifest uses the XML add-in format to support Mac as well as Windows and web Outlook. Its Office permission is **ReadWriteItem**: the ability to update the current message and attach the two small logo images. The code changes only the signature slot; it never sends messages. An administrator must review the real consent screens before installation.

Check a representative mailbox on each of Mac, Windows and web Outlook. In a message, open **ARK signatures**. The panel tries Outlook SSO automatically. If Microsoft requires consent or an authentication check, choose **Continue with Microsoft 365**. The preview should contain the correct profile. Event code uses silent authentication; it cannot display sign-in popups itself.

Microsoft may list the standard delegated sign-in scopes `openid`, `profile` and `offline_access` alongside `User.Read`. Organisation-wide consent authorises each employee to use the app under their own identity; it does not grant access to other employees' profiles. The code retrieves `/me`, not the directory user collection.

When ready to test automatic insertion, set `enabled` to `true` and publish. The build refuses to enable insertion without configured tenant and client IDs. The add-in only runs for users to whom it has been assigned.

## 5. Manage employee rollout

The owner has accepted the general release. Review the employee list before expanding the deployment group, and continue [ACCEPTANCE.md](ACCEPTANCE.md) with real send/receive tests. Keep a documented local signature fallback until availability and client coverage are established. The Microsoft 365 deployment confirmation advises allowing up to 72 hours for the add-in to appear and says users may need to relaunch Microsoft 365.

To add another employee to the existing installation, open **Integrated apps → ARK signatures → Users**, choose **Specific users or groups**, retain the current assigned users, add the new account, and save **Update**. Verify the saved list before announcing availability. Do not upload another copy of the same add-in. Keep named deployment membership in an internal record rather than this public repository.

Employees do not need to install the add-in or copy HTML themselves. With the assignment delivered, a supported Outlook client, `enabled:true`, and a usable Microsoft 365 session, compose events insert signatures automatically. If silent authentication cannot complete, the employee must open **ARK signatures → Continue with Microsoft 365** once to complete Microsoft's required check; background events cannot show a sign-in prompt. The panel should then show their own profile. **Refresh this message** updates an existing draft; new events use the central template. Verify any existing personal signature does not create duplicates during each employee's first test.

For routine wording or branding updates, edit `branding.json` or the templates and publish. Administrator reinstall is normally unnecessary for content changes. Manifest changes — new URLs, IDs, permissions, events or requirements — need a manifest version bump and an administrator-managed update.

## Pause and rollback

- To pause automatic insertion, set `enabled:false` and publish. After the new bundle becomes available, new compose events retain the existing client signature.
- To roll back a design, revert its GitHub commit and republish.
- To remove the deployment, unassign or remove the add-in in Integrated apps and restore the usual Outlook signature defaults.
- An already open draft may contain a previously inserted signature; pausing doesn't edit existing drafts or sent emails.

## Release and deployment status

The single-tenant **ARK central signatures** app is registered. Its identifiers are configured in `deployment.json`, and its only requested Graph permission is delegated `User.Read`. No client secret is used.

The installation began on 8 September 2026 and was subsequently expanded to three named users. The first user confirmed availability on Mac, web and new Windows Outlook, and approved received full and compact forward signatures. On 9 September 2026, the owner confirmed that a second user's add-in appeared automatically after the deployment delay and accepted the transition to general release. The third user's availability remains unconfirmed.

The wider employee assignment is a separate administrative action. The owner requested review of the proposed employee list before it is applied. Keep employee names, contact details and assignment evidence in internal records, outside this public repository.

Version 1.0.2 retries a completed, recoverable token-acquisition failure once silently before showing the compose warning. It shares a two-attempt limit with Graph 401 recovery and allows at most 60 seconds for the overall automatic event, with shorter per-request deadlines. Background failures emit safe support references to the add-in console; pane failures expose references under **Support details**. These paths have separate runtime state, so a successful pane preview does not display the prior background exception. No tokens or employee profiles are included in those diagnostics.

If **Retry connection** cannot clear a Microsoft broker sign-in rejection, fully quit and reopen Outlook. This improves recovery and diagnosis; it does not establish a permanent Microsoft-side fix. The Outlook manifest remains version `1.0.0.0`; existing users do not need a new manifest for this release. Outlook may cache hosted code in an already-open session. Close existing drafts and reopen the add-in, or restart Outlook once, to establish the new baseline before testing an extended idle session.
