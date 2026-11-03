# Isadek Office for organisations

## Installing

- **Windows**: `Isadek-Office-x.y.z-x64.msi` installs for every user of a computer and can be
  deployed with Intune, Configuration Manager or Group Policy (`msiexec /i Isadek-Office-x.y.z-x64.msi /qn`).
  MSI installs don't update themselves: deploy the next MSI. The per-user installers
  (`Setup-…-x64.exe`, `-arm64.exe`) and the portable builds are also available.
- **macOS**: the universal `.dmg` or `.zip` (Intel and Apple Silicon).
- **Linux**: AppImage, `.deb` or `.tar.gz`, each for x64 and ARM64.

## Policies

Policies turn features off on managed computers. The desktop app reads them when it starts, from
places only an administrator can change (later ones win when the same policy is set twice):

| Where | How |
| --- | --- |
| A JSON file | Windows `%ProgramData%\Isadek Office\policies.json`, macOS `/Library/Application Support/Isadek Office/policies.json`, Linux `/etc/isadek-office/policies.json` |
| Windows Group Policy / Intune | the ADMX template in [`policies/`](policies/) (copy `isadek-office.admx` to `PolicyDefinitions` and `en-US/isadek-office.adml` beside it); values go under `HKLM\SOFTWARE\Policies\Isadek\Office` |
| macOS configuration profile | preference domain `io.github.isaadac.isadekoffice` (read from `/Library/Managed Preferences`) |

| Policy | Value | Effect |
| --- | --- | --- |
| `DisableAssistant` | true/false (1/0) | No Assistant (Claude) button; its requests are refused |
| `DisableCloudStorage` | true/false | No Google Drive, OneDrive, Dropbox or WebDAV |
| `DisableLiveEditing` | true/false | No live sessions |
| `LiveServers` | list of `ws://`/`wss://` addresses (JSON array, or separated by `;`) | Live sessions only on these servers; include `local` to allow sessions hosted on computers in the local network (private addresses) |
| `DisableForms` | true/false | No publishing forms, live polls, audience Q&A or the phone remote |
| `DisableAddins` | true/false | No installing or running add-ins |
| `DisableWebData` | true/false | No From Web, Web Clip or other downloads a page asks for |
| `DisableAutoUpdate` | true/false | No automatic updates; the menu option is locked off |
| `DisableCrashReports` | true/false | Crash reports are never sent; the menu option is locked off |
| `DisableSpeechDownload` | true/false | On-device speech recognition can't download its model |
| `DefaultLanguage` | `en`, `es`, `fr`, `de`, `pt`, `it` or `ar` | The interface language until a person chooses one |

Example `policies.json`:

```json
{
  "DisableAssistant": true,
  "DisableCloudStorage": true,
  "LiveServers": ["wss://live.example.org/", "local"],
  "DisableAutoUpdate": true,
  "DefaultLanguage": "fr"
}
```

What a policy turns off is refused in the app's main process, not only hidden in the windows, so
a document or an add-in can't get round it: its functions are refused there, and connections to
live-editing servers, form servers and the speech model download are checked as they leave the app. Unknown names and values of the wrong kind are ignored.
Policies apply to the desktop app; the web version is configured on its server (see the README).
