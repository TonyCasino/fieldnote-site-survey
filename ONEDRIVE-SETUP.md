# Connect Fieldnote to Microsoft 365

Fieldnote is set up for people to sign in with their normal Microsoft work account. An administrator needs to create one Microsoft Entra application before that can be enabled.

1. In the Microsoft Entra admin center, open **App registrations** and choose **New registration**.
2. Name it `Fieldnote Site Survey` and choose **Accounts in this organizational directory only**.
3. Under **Authentication**, add a **Single-page application** redirect URI for the app address (for example, `https://fieldnote-site-survey.example.com`). Add the final production address too, if it differs.
4. Under **API permissions**, add delegated Microsoft Graph permissions: `User.Read`, `Files.ReadWrite`, and `offline_access`. Grant administrator consent if your organization requires it.
5. Copy the **Application (client) ID** from the Overview screen into `NEXT_PUBLIC_MICROSOFT_CLIENT_ID` in a local `.env` file, based on `.env.example`.

Do not create a client secret for this browser application. Each employee signs in directly with Microsoft, and files are uploaded only to the OneDrive account or shared folder that they choose.

The later upload connection will create a site folder, create station subfolders, upload the original photo, and preserve any edited copy beside it. Microsoft Graph supports the sign-in and OneDrive file APIs used for that flow.
