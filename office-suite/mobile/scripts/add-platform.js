'use strict';
// Adds the native project once (npx cap add), then applies Isadek Office's settings to it:
// on Android, open .docx/.xlsx/.pptx and friends from other apps ("Open with").
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

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
