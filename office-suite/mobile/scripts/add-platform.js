'use strict';
// Adds the native project once (npx cap add), then applies Isadek Office's settings to it:
// "Open with" for .docx/.xlsx/.pptx and friends, the print/PDF plugin (native/), and the
// app icon and splash screen made from build/icon.png.
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

/** app/build.gradle with the app's version and a release signing config read from the environment. */
function androidGradle(src, version) {
  const [a, b, c] = String(version).split(/[.-]/).map(n => parseInt(n, 10) || 0);
  const code = process.env.ANDROID_VERSION_CODE ? parseInt(process.env.ANDROID_VERSION_CODE, 10) : a * 10000 + b * 100 + c;
  let out = src.replace(/versionCode \d+/, `versionCode ${code}`).replace(/versionName "[^"]*"/, `versionName "${version}"`);
  if (!out.includes('isadek-release-signing')) out += `
// isadek-release-signing: release builds are signed with the keystore named in the environment.
android {
    signingConfigs {
        release {
            if (System.getenv("ANDROID_KEYSTORE_FILE")) {
                storeFile file(System.getenv("ANDROID_KEYSTORE_FILE"))
                storePassword System.getenv("ANDROID_KEYSTORE_PASSWORD")
                keyAlias System.getenv("ANDROID_KEY_ALIAS")
                keyPassword System.getenv("ANDROID_KEY_PASSWORD") ?: System.getenv("ANDROID_KEYSTORE_PASSWORD")
            }
        }
    }
    buildTypes {
        release {
            if (System.getenv("ANDROID_KEYSTORE_FILE")) signingConfig signingConfigs.release
        }
    }
}
`;
  return out;
}
if (require.main !== module) { module.exports = { androidGradle }; return; }

const platform = process.argv[2];
if (!['android', 'ios'].includes(platform)) { console.error('Usage: add-platform.js android|ios'); process.exit(1); }
const dir = path.resolve(__dirname, '..', platform);
if (!fs.existsSync(dir)) execSync(`npx cap add ${platform}`, { cwd: path.resolve(__dirname, '..'), stdio: 'inherit' });

const MIME = [
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.ms-excel', 'application/vnd.oasis.opendocument.spreadsheet',
  'text/csv', 'text/comma-separated-values', 'text/tab-separated-values', 'text/markdown', 'text/plain', 'text/html'
];
if (platform === 'android') {
  const manifest = path.join(dir, 'app/src/main/AndroidManifest.xml');
  let xml = fs.readFileSync(manifest, 'utf8');
  if (!xml.includes('atlas-open-with')) {
    const filters = `
            <!-- atlas-open-with: offer Isadek Office for office documents -->
            <intent-filter>
                <action android:name="android.intent.action.VIEW" />
                <category android:name="android.intent.category.DEFAULT" />
                <category android:name="android.intent.category.BROWSABLE" />
                <data android:scheme="content" />
                <data android:scheme="file" />
${MIME.map(m => `                <data android:mimeType="${m}" />`).join('\n')}
            </intent-filter>`;
    // Inside MainActivity, after its launcher intent-filter.
    xml = xml.replace(/(<activity[\s\S]*?android:name="[^"]*MainActivity"[\s\S]*?<\/intent-filter>)/, `$1${filters}`);
    if (!xml.includes('WRITE_EXTERNAL_STORAGE')) {
      xml = xml.replace(/(<uses-permission android:name="android.permission.INTERNET" \/>)/, '$1\n    <uses-permission android:name="android.permission.READ_EXTERNAL_STORAGE" android:maxSdkVersion="32" />\n    <uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" android:maxSdkVersion="29" />');
    }
    fs.writeFileSync(manifest, xml);
    console.log('AndroidManifest.xml: added "Open with" for office documents');
  }
}
if (platform === 'ios') {
  const plist = path.join(dir, 'App/App/Info.plist');
  let xml = fs.readFileSync(plist, 'utf8');
  if (!xml.includes('LSSupportsOpeningDocumentsInPlace')) {
    const types = ['org.openxmlformats.wordprocessingml.document', 'org.openxmlformats.spreadsheetml.sheet', 'org.openxmlformats.presentationml.presentation', 'public.comma-separated-values-text', 'public.plain-text', 'net.daringfireball.markdown'];
    xml = xml.replace(/<dict>/, `<dict>
	<key>LSSupportsOpeningDocumentsInPlace</key>
	<false/>
	<key>UIFileSharingEnabled</key>
	<true/>
	<key>CFBundleDocumentTypes</key>
	<array>
		<dict>
			<key>CFBundleTypeName</key>
			<string>Office document</string>
			<key>LSHandlerRank</key>
			<string>Alternate</string>
			<key>LSItemContentTypes</key>
			<array>
${types.map(t => `				<string>${t}</string>`).join('\n')}
			</array>
		</dict>
	</array>`);
    fs.writeFileSync(plist, xml);
    console.log('Info.plist: added document types');
  }
}

// ---------- Version and release signing (Android) ----------
// The version comes from the desktop app's package.json (1.2.3 → versionName 1.2.3, versionCode
// 10203). Release builds are signed when ANDROID_KEYSTORE_FILE (and its passwords) are set;
// without them only debug builds can be installed.
if (platform === 'android') {
  const gradle = path.join(dir, 'app/build.gradle');
  fs.writeFileSync(gradle, androidGradle(fs.readFileSync(gradle, 'utf8'), require('../../package.json').version));
  console.log('build.gradle: version and release signing');
}

// ---------- Printing and PDF (native/) ----------
const NATIVE = path.resolve(__dirname, '../native');
if (platform === 'android') {
  const javaRoot = path.join(dir, 'app/src/main/java');
  const main = fs.readdirSync(javaRoot, { recursive: true }).map(f => path.join(javaRoot, f)).find(f => f.endsWith(path.sep + 'MainActivity.java'));
  let src = fs.readFileSync(main, 'utf8');
  const pkg = src.match(/^package ([\w.]+);/m)[1];
  fs.writeFileSync(path.join(path.dirname(main), 'PrintPlugin.java'), fs.readFileSync(path.join(NATIVE, 'android/PrintPlugin.java'), 'utf8').replace('__PACKAGE__', pkg));
  if (!src.includes('PrintPlugin')) {
    src = src.replace(/public class MainActivity extends BridgeActivity \{\s*\}/, `public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(android.os.Bundle savedInstanceState) {
        registerPlugin(PrintPlugin.class);
        super.onCreate(savedInstanceState);
    }
}`);
    if (!src.includes('PrintPlugin')) { console.error('MainActivity.java: could not register the print plugin'); process.exit(1); }
    fs.writeFileSync(main, src);
    console.log('MainActivity.java: registered the print plugin');
  }
}
if (platform === 'ios') {
  // In AppDelegate.swift, which is already part of the Xcode target.
  const delegate = path.join(dir, 'App/App/AppDelegate.swift');
  const swift = fs.readFileSync(delegate, 'utf8');
  if (!swift.includes('isadek-print')) fs.writeFileSync(delegate, swift + fs.readFileSync(path.join(NATIVE, 'ios/IsadekPrint.swift'), 'utf8'));
  const board = path.join(dir, 'App/App/Base.lproj/Main.storyboard');
  const xml = fs.readFileSync(board, 'utf8');
  if (!xml.includes('IsadekViewController')) {
    const out = xml.replace(/customClass="CAPBridgeViewController" customModule="Capacitor"/, 'customClass="IsadekViewController" customModule="App" customModuleProvider="target"');
    if (out === xml) { console.error('Main.storyboard: could not use the Isadek view controller'); process.exit(1); }
    fs.writeFileSync(board, out);
    console.log('Main.storyboard: uses the Isadek view controller (print plugin)');
  }
}

// ---------- Icon and splash screen ----------
const assets = path.resolve(__dirname, '../.assets');
fs.mkdirSync(assets, { recursive: true });
fs.copyFileSync(path.resolve(__dirname, '../../build/icon.png'), path.join(assets, 'logo.png'));
execSync(`npx capacitor-assets generate --${platform} --assetPath .assets --iconBackgroundColor "#ffffff" --iconBackgroundColorDark "#1f1f1f" --splashBackgroundColor "#f3f2f1" --splashBackgroundColorDark "#1f1f1f"`, { cwd: path.resolve(__dirname, '..'), stdio: 'inherit' });
