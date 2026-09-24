# Fieldnote iOS app

Fieldnote's production web app and Azure API are the shared service layer for the iOS app. The native app should preserve the existing station workflow while replacing browser-specific authentication, camera access, and local file storage with iOS APIs.

## Native application plan

1. Create the Xcode project on a Mac with a permanent bundle identifier owned by the publishing organization.
2. Add camera and photo-library usage descriptions to `Info.plist`.
3. Use MSAL for iOS with the existing Microsoft Entra client ID `ced34bd9-a501-4fad-8ba4-7bb5db371c82` and tenant ID `513f9529-ec84-415b-a0b6-e97302b1524f`.
4. Add the iOS/macOS platform in Microsoft Entra for the chosen bundle identifier. The redirect URI will follow `msauth.<bundle-id>://auth`.
5. Store project metadata locally and photos in the app's Application Support directory. Keep a OneDrive copy through the existing Azure upload API.
6. Load existing projects from the **Fieldnote Saved Projects** folder in OneDrive so users can move between the web and native apps.
7. Test on a physical iPhone, distribute through TestFlight, then prepare App Store privacy details and screenshots.

## Data migration

The current Home Screen app and a native iOS app have separate local storage. Upload each existing project to OneDrive before removing the Home Screen app. The native app can then restore it from **Fieldnote Saved Projects** after Microsoft sign-in.

## Mac build requirements

- A current Mac and Xcode release
- An Apple Developer account belonging to the publishing individual or organization
- A permanent bundle identifier
- Microsoft Entra permission to add the iOS/macOS platform and redirect URI

Do not put Apple signing certificates, provisioning profiles, Microsoft tokens, or client secrets in this repository.
