#!/usr/bin/env bash
# Build/refresh only the Mac menu app; never starts or stops the product stack.
# Usage: bash scripts/macos-app.sh [installation-folder] [application-path]
set -euo pipefail
source_scripts="$(cd "$(dirname "$0")" && pwd)"
repository="$(cd "${1:-$source_scripts/..}" && pwd)"
app="${2:-$HOME/Applications/doc.haus.app}"
if [ ! -f "$repository/start.sh" ] || [ ! -f "$repository/scripts/data-paths.sh" ] || [ ! -f "$repository/scripts/start-runtime.sh" ]; then
  echo "doc.haus: choose an installation containing start.sh and its runtime helpers." >&2
  exit 1
fi
if [ -L "$app" ] || { [ -e "$app" ] && [ ! -d "$app" ]; }; then
  echo "doc.haus: the application path is not a regular app folder." >&2
  exit 1
fi
mkdir -p "$(dirname "$app")"
stage=$(mktemp -d "$(dirname "$app")/.doc.haus-build.XXXXXX")
trap 'rm -rf "$stage"' EXIT
candidate="$stage/doc.haus.app"
osacompile -s -o "$candidate" "$source_scripts/macos-app.applescript"
resources="$candidate/Contents/Resources"
mkdir -p "$resources"
{
  printf '#!/bin/bash\nexport DOCHAUS_INSTALL_DIR=%q\n' "$repository"
  # Only installation paths and ports survive Finder's separate environment.
  # Never copy provider credentials or the caller's general environment.
  for variable in DOCHAUS_DATA_ROOT WORKSPACE_ROOT DOCHAUS_RUNTIME_DIR DOC_HAUS_ENGINE_PORT DOC_HAUS_INGEST_PORT DOC_HAUS_WEB_PORT; do
    if [ -n "${!variable:-}" ]; then printf 'export %s=%q\n' "$variable" "${!variable}"; fi
  done
} > "$resources/installation.sh"
cp "$source_scripts/macos-app-launch.sh" "$resources/open-doc-haus.sh"
cat > "$resources/doc.haus.command" <<'COMMAND'
#!/bin/bash
set -euo pipefail
. "$(dirname "$0")/installation.sh"
exec "$DOCHAUS_INSTALL_DIR/start.sh"
COMMAND
chmod +x "$resources/doc.haus.command" "$resources/open-doc-haus.sh"
icon="$source_scripts/icons/doc-haus.icns"
if [ ! -f "$icon" ]; then icon="$repository/scripts/icons/doc-haus.icns"; fi
if [ -f "$icon" ]; then cp "$icon" "$resources/AppIcon.icns"; fi
plist="$candidate/Contents/Info.plist"
plutil -replace CFBundleIdentifier -string haus.doc.start "$plist"
plutil -replace CFBundleName -string doc.haus "$plist"
plutil -replace CFBundleDisplayName -string doc.haus "$plist"
plutil -replace CFBundleIconFile -string AppIcon "$plist"
# Hide the temporary launcher from the Dock; the branded Applications shortcut
# remains available while the browser and Terminal own the running interface.
plutil -replace LSUIElement -bool true "$plist"
plutil -replace NSAppleEventsUsageDescription -string 'Find and focus your existing doc.haus tab in your default browser. Only tab addresses are checked; page contents are not read or saved.' "$plist"
# osacompile supplies a generic icon name and unrelated permission descriptions.
# This wrapper uses its ICNS file and only automates the selected browser.
for key in CFBundleIconName NSAppleMusicUsageDescription NSCalendarsUsageDescription NSCameraUsageDescription NSContactsUsageDescription NSHomeKitUsageDescription NSMicrophoneUsageDescription NSPhotoLibraryUsageDescription NSRemindersUsageDescription NSSiriUsageDescription NSSystemAdministrationUsageDescription; do
  if plutil -extract "$key" raw -o /dev/null "$plist" >/dev/null 2>&1; then plutil -remove "$key" "$plist"; fi
done
# Seal the locally generated app after setting its identity and resources.
codesign --force --sign - "$candidate"

# Compile and sign before touching the installed entry. Keep a recovery copy;
# if the final move fails, restore it without touching any runtime or user data.
backup=
if [ -e "$app" ]; then
  backup="$app.previous-$(date +%Y%m%d%H%M%S)-$$"
  mv "$app" "$backup"
fi
if ! mv "$candidate" "$app"; then
  if [ -n "$backup" ]; then mv "$backup" "$app"; fi
  exit 1
fi
touch "$app"
echo "doc.haus: Mac app ready at $app"
if [ -n "$backup" ]; then echo "doc.haus: previous app kept at $backup"; fi
