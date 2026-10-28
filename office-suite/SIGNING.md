# Signing the installers

Unsigned installers work, but Windows SmartScreen and macOS Gatekeeper warn about them, and on macOS
auto-update only works for signed apps. The **Isadek Office** workflow signs automatically as soon as you
add these repository secrets (GitHub → Settings → Secrets and variables → Actions → New repository secret).
Without them it builds unsigned, as before.

## Windows

| Secret | Value |
|---|---|
| `WIN_CSC_LINK` | Your code-signing certificate (`.pfx`/`.p12`) encoded as base64: `base64 -w0 cert.pfx` (macOS: `base64 -i cert.pfx`) |
| `WIN_CSC_KEY_PASSWORD` | The certificate's password |

Any OV or EV code-signing certificate from a public CA works. An EV certificate removes SmartScreen
warnings immediately; an OV certificate builds reputation over time. EV certificates on a hardware token
can't be used from CI as-is; use your CA's cloud signing service instead.

If you use the publisher name on the certificate, update `build.win.signtoolOptions.publisherName` in
`package.json` to match it exactly.

## macOS

You need a paid Apple Developer account ($99/year).

| Secret | Value |
|---|---|
| `MAC_CSC_LINK` | Your **Developer ID Application** certificate exported from Keychain as `.p12`, base64-encoded |
| `MAC_CSC_KEY_PASSWORD` | The `.p12` export password |
| `APPLE_ID` | The Apple ID email of the developer account (for notarization) |
| `APPLE_APP_SPECIFIC_PASSWORD` | An app-specific password from appleid.apple.com → Sign-In and Security |
| `APPLE_TEAM_ID` | Your 10-character Team ID from developer.apple.com → Membership |

With all five secrets the app is signed **and notarized**, so it opens without any warning.

## Linux

AppImage and `.deb` packages don't need signing.

## Checking

Open the latest **Isadek Office** run in the Actions tab. An unsigned build shows the notice
"Building unsigned installers". A signed macOS build logs "Notarizing with Apple".
