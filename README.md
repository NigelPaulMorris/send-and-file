# Send & File for the new Outlook

A SimplyFile-style add-in. When you press **Send**, Outlook asks where the email should be filed. You pick a folder, press Send again, and the email is moved out of Sent Items into that folder.

## How it behaves

1. You press Send on an email you haven't chosen a folder for.
2. Outlook shows a prompt with two buttons. **Choose folder** opens the Send & File pane. **Send Anyway** sends it and leaves it in Sent Items.
3. In the pane, type part of a folder name or click a suggestion. Suggestions come from where the rest of the conversation is filed and where you last filed mail to the same people. **Don't file** is there too.
4. Press Send again. It goes straight through.
5. The sent copy is moved to your folder shortly afterwards, as described below.

You can also open the pane *before* sending, using the **File when sent** button in the compose toolbar, so you never see the prompt.

When reading an email, the **File email** button moves the open email to a folder, which is the classic SimplyFile one-click filing.

### Why the move happens "shortly after" rather than instantly

New Outlook add-ins aren't allowed to change where Outlook saves a sent email. So Send & File writes your choice onto the email itself, and that tag stays on the copy in Sent Items. The add-in then moves tagged emails out of Sent Items whenever:

- you start a new email or reply, which catches everything sent before it,
- the Send & File pane is open, which checks every minute.

**Tip:** open any email, click **File email**, and click the pin icon on the pane. While it's pinned, sent mail is filed within about a minute.

---

## Setup (about 20 minutes, one time)

You need three things: somewhere to host the files, an app registration so the add-in can see your folders, and to install the manifest in Outlook.

### 1. Host the files (free, using GitHub Pages)

1. Create a free GitHub account if you don't have one, then create a **new public repository** named `send-and-file`.
2. Upload every file in this folder to it, including the `assets` folder. You don't need to upload the README.
3. In the repository, go to **Settings → Pages**, set *Source* to **Deploy from a branch**, pick `main` and `/ (root)`, and save.
4. After a minute your files are live at `https://YOURNAME.github.io/send-and-file/`. Check that `https://YOURNAME.github.io/send-and-file/taskpane.html` opens.

Any HTTPS web host works. GitHub Pages is just the easiest free option.

### 2. Register the app in Microsoft Entra

This lets the add-in use Outlook's own sign-in to list your folders and move emails.

1. Go to <https://entra.microsoft.com> → **Identity → Applications → App registrations → New registration**.
   - For a personal Outlook.com account, use <https://portal.azure.com> → *App registrations*.
2. Name it `Send & File`.
3. Under *Supported account types*, choose **Accounts in any organizational directory and personal Microsoft accounts**. If you only use a work account, **single tenant** is fine.
4. Under *Redirect URI*, choose **Single-page application (SPA)** and enter:
   `brk-multihub://YOURNAME.github.io`
   This is the domain only, with no `/send-and-file` path and no trailing slash.
5. Click **Register** and copy the **Application (client) ID**.
6. Go to **API permissions → Add a permission → Microsoft Graph → Delegated** and add **Mail.ReadWrite**.
   If your organisation doesn't let users approve apps themselves, ask IT to click **Grant admin consent**.

### 3. Fill in the two placeholders

- In `shared.js`, replace `YOUR-CLIENT-ID` with the client ID you just copied.
- In `manifest.xml`, find and replace every `https://YOUR-HOST` with `https://YOURNAME.github.io/send-and-file`.

Re-upload `shared.js` to GitHub after editing it. The manifest stays on your computer.

### 4. Install it in the new Outlook

1. In a browser, go to <https://aka.ms/olksideload>. This opens Outlook on the web at the add-ins dialog.
2. Choose **My add-ins → Custom add-ins → Add a custom add-in → Add from file**, then select your `manifest.xml`.
3. Restart the new Outlook. The add-in appears there automatically, because it's tied to your mailbox.

If **Add a custom add-in** is missing, your organisation has turned off custom add-ins. An admin can deploy `manifest.xml` for you in the Microsoft 365 admin center under **Settings → Integrated apps → Upload custom apps**.

### 5. First run

Open any email and click **File email**. The first time, you'll be asked to sign in and approve access. After that, try sending an email.

---

## Testing checklist

1. Send an email to yourself without picking a folder. You should get the prompt.
2. Click **Choose folder**, pick a folder, and press Send again.
3. Start a new email, or have the pinned pane open. Within a minute the sent copy should leave Sent Items and appear in your folder.

## Troubleshooting

- **No prompt when sending.** Confirm the add-in is installed and restart Outlook. Prompts on send need an up-to-date new Outlook.
- **Folders never load / sign-in fails.** Check the redirect URI is exactly `brk-multihub://YOURNAME.github.io`, the type is *Single-page application*, and `shared.js` has your client ID.
- **Email stays in Sent Items.** Open the pane once, which triggers a check. If it still doesn't move, press F12 in Outlook on the web and look for "Send & File" messages in the console.
- **Prompt appears, but Send Anyway is what you want for most emails.** Use the **File when sent** button before sending when you do want to file, and Send Anyway otherwise.

## Files

| File | What it does |
|---|---|
| `manifest.xml` | Tells Outlook about the add-in, its buttons, and the on-send prompt |
| `taskpane.html` / `taskpane.js` | The folder picker pane |
| `commands.html` / `launchevent.js` | The Send check and the background filing on new compose |
| `shared.js` | Settings, sign-in, folder loading and filing |
| `msal-browser.min.js` | Microsoft's sign-in library (v3.30.0), included so nothing loads from third parties |
| `assets/` | Icons |
