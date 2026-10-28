# Signing, stores and release settings

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

## Android (release APK and Google Play)

Without a keystore the workflow builds a debug APK. With one it builds a signed release APK (attached to
GitHub Releases) and an Android App Bundle for Google Play.

| Secret | Value |
|---|---|
| `ANDROID_KEYSTORE` | The upload keystore, base64-encoded. Create one with `keytool -genkeypair -v -keystore isadek.jks -keyalg RSA -keysize 4096 -validity 10000 -alias isadek` |
| `ANDROID_KEYSTORE_PASSWORD` | Its password |
| `ANDROID_KEY_ALIAS` | The key's alias (`isadek` above) |
| `ANDROID_KEY_PASSWORD` | The key's password, if different from the keystore's |
| `PLAY_SERVICE_ACCOUNT_JSON` | Optional: a Google Play service account's JSON key. On `office-v*` tags the bundle goes to the internal testing track |

Keep a backup of the keystore: Google Play only accepts updates signed with the same key (or enrol in Play App
Signing, recommended). The version code is `major×10000 + minor×100 + patch` of `package.json`'s version;
set the repository variable `ANDROID_VERSION_CODE` to override it.

## Stores

Each store build runs only when its credentials are set up; on `office-v*` tags it is also uploaded.

### Microsoft Store

An MSIX package (`*.appx`, artifact `store-Windows`) is built on every run. Reserve the app name in
[Partner Center](https://partner.microsoft.com/dashboard), then set these **repository variables** (Settings →
Secrets and variables → Actions → Variables) from Product identity:

| Variable | Value |
|---|---|
| `ISADEK_APPX_IDENTITY` | Package/Identity/Name, e.g. `12345Isadek.IsadekOffice` |
| `ISADEK_APPX_PUBLISHER` | Package/Identity/Publisher, e.g. `CN=1A2B3C4D-…` |
| `ISADEK_APPX_PUBLISHER_NAME` | Package/Properties/PublisherDisplayName |

Upload the `.appx` in a Partner Center submission; the Store signs it. The Store copy updates through the Store.

### Mac App Store

Needs the Apple Developer account above, an App Store Connect app record for `io.github.isaadac.isadekoffice`,
and a Mac App Store provisioning profile.

| Secret | Value |
|---|---|
| `MAS_CSC_LINK` / `MAS_CSC_KEY_PASSWORD` | **Apple Distribution** (or 3rd Party Mac Developer Application) certificate, `.p12` base64 and password |
| `MAS_INSTALLER_CSC_LINK` / `MAS_INSTALLER_CSC_KEY_PASSWORD` | **3rd Party Mac Developer Installer** certificate, `.p12` base64 and password |
| `MAS_PROVISIONING_PROFILE` | The `.provisionprofile`, base64 |
| `APP_STORE_CONNECT_KEY_ID`, `APP_STORE_CONNECT_ISSUER_ID`, `APP_STORE_CONNECT_KEY` | An App Store Connect API key (Users and Access → Integrations): its ID, the issuer ID and the `.p8` file's contents. Used for uploading |

The app runs in the sandbox there (`build/entitlements.mas.plist`): files the person opens or saves, the network,
and the microphone for narration and audio notes.

### App Store (iPhone and iPad)

| Secret | Value |
|---|---|
| `IOS_DIST_CERT_P12` / `IOS_DIST_CERT_PASSWORD` | **Apple Distribution** certificate, `.p12` base64 and password |
| `IOS_PROVISIONING_PROFILE` | An App Store provisioning profile for `io.github.isaadac.isadekoffice`, base64 |
| `APPLE_TEAM_ID` and the three `APP_STORE_CONNECT_*` secrets | As above |

The workflow then builds `Isadek-Office-<version>-ios.ipa` (`mobile/scripts/ios-release.sh`) and, on tags,
uploads it to App Store Connect for TestFlight and review.

### Snap Store

Register the name `isadek-office` at [snapcraft.io](https://snapcraft.io), then export credentials with
`snapcraft export-login --snaps=isadek-office --acls package_access,package_push,package_release -` and save
them as the secret `SNAPCRAFT_STORE_CREDENTIALS`. Tags release to the stable channel.

## Crash reports and the download site

| Variable | Value |
|---|---|
| `ISADEK_CRASH_URL` | Where the app sends crash reports when people turn them on (https). The bundled web server collects them at `/api/crash` when started with `CRASH_DIR=/some/folder`. Without it, reports stay on the computer |
| `ISADEK_PAGES` | `true` to publish the download site to GitHub Pages on each release (first set Settings → Pages → Source to **GitHub Actions**) |

## Checking

Open the latest **Isadek Office** run in the Actions tab. An unsigned build shows the notice
"Building unsigned installers". A signed macOS build logs "Notarizing with Apple". An Android build without a
keystore shows "No release keystore". Store steps that are skipped had no credentials.

Each release has `SHA256SUMS.txt`; check a download with `sha256sum -c SHA256SUMS.txt --ignore-missing`.
