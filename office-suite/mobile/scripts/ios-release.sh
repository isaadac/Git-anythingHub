#!/usr/bin/env bash
# Builds a signed iOS app for the App Store (run on macOS after `npm run ios`), and uploads it
# when UPLOAD=1. Needs (from the repository secrets, see SIGNING.md):
#   IOS_DIST_CERT_P12 (base64), IOS_DIST_CERT_PASSWORD, IOS_PROVISIONING_PROFILE (base64),
#   APPLE_TEAM_ID; for uploading also APP_STORE_CONNECT_KEY_ID, APP_STORE_CONNECT_ISSUER_ID,
#   APP_STORE_CONNECT_KEY (the .p8 file's contents).
set -euo pipefail
cd "$(dirname "$0")/../ios/App"
TMP="${RUNNER_TEMP:-$(mktemp -d)}"
VERSION=$(node -p "require('../../../package.json').version")
BUILD="${IOS_BUILD_NUMBER:-${GITHUB_RUN_NUMBER:-1}}"

# The distribution certificate, in a keychain of its own.
KEYCHAIN="$TMP/isadek.keychain-db"
KCPASS=$(openssl rand -hex 16)
echo "$IOS_DIST_CERT_P12" | base64 --decode > "$TMP/dist.p12"
security create-keychain -p "$KCPASS" "$KEYCHAIN"
security set-keychain-settings -lut 3600 "$KEYCHAIN"
security unlock-keychain -p "$KCPASS" "$KEYCHAIN"
security import "$TMP/dist.p12" -k "$KEYCHAIN" -P "$IOS_DIST_CERT_PASSWORD" -T /usr/bin/codesign -T /usr/bin/security
security set-key-partition-list -S apple-tool:,apple: -s -k "$KCPASS" "$KEYCHAIN" > /dev/null
security list-keychains -d user -s "$KEYCHAIN" $(security list-keychains -d user | tr -d '"')
rm "$TMP/dist.p12"

# The provisioning profile.
echo "$IOS_PROVISIONING_PROFILE" | base64 --decode > "$TMP/app.mobileprovision"
security cms -D -i "$TMP/app.mobileprovision" > "$TMP/profile.plist"
UUID=$(/usr/libexec/PlistBuddy -c 'Print UUID' "$TMP/profile.plist")
PROFILE_NAME=$(/usr/libexec/PlistBuddy -c 'Print Name' "$TMP/profile.plist")
mkdir -p "$HOME/Library/MobileDevice/Provisioning Profiles"
cp "$TMP/app.mobileprovision" "$HOME/Library/MobileDevice/Provisioning Profiles/$UUID.mobileprovision"
BUNDLE=$(node -p "require('../../capacitor.config.json').appId")

xcodebuild -project App.xcodeproj -scheme App -configuration Release -sdk iphoneos \
  -archivePath "$TMP/App.xcarchive" -destination 'generic/platform=iOS' archive \
  MARKETING_VERSION="$VERSION" CURRENT_PROJECT_VERSION="$BUILD" \
  CODE_SIGN_STYLE=Manual DEVELOPMENT_TEAM="$APPLE_TEAM_ID" CODE_SIGN_IDENTITY="Apple Distribution" \
  PROVISIONING_PROFILE_SPECIFIER="$PROFILE_NAME" OTHER_CODE_SIGN_FLAGS="--keychain $KEYCHAIN"

cat > "$TMP/ExportOptions.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>method</key><string>app-store-connect</string>
  <key>teamID</key><string>$APPLE_TEAM_ID</string>
  <key>signingStyle</key><string>manual</string>
  <key>provisioningProfiles</key><dict><key>$BUNDLE</key><string>$PROFILE_NAME</string></dict>
</dict></plist>
PLIST
xcodebuild -exportArchive -archivePath "$TMP/App.xcarchive" -exportOptionsPlist "$TMP/ExportOptions.plist" -exportPath "$TMP/export"
OUT="${GITHUB_WORKSPACE:-$PWD}/Isadek-Office-$VERSION-ios.ipa"
mv "$TMP"/export/*.ipa "$OUT"
echo "Built $OUT"

if [ "${UPLOAD:-}" = "1" ]; then
  mkdir -p "$HOME/.appstoreconnect/private_keys"
  printf '%s' "$APP_STORE_CONNECT_KEY" > "$HOME/.appstoreconnect/private_keys/AuthKey_$APP_STORE_CONNECT_KEY_ID.p8"
  xcrun altool --upload-app -f "$OUT" -t ios --apiKey "$APP_STORE_CONNECT_KEY_ID" --apiIssuer "$APP_STORE_CONNECT_ISSUER_ID"
fi
security delete-keychain "$KEYCHAIN" || true
