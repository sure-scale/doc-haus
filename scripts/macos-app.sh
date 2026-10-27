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
. "$repository/scripts/data-paths.sh"
configure_data_paths "$repository"
DOC_HAUS_ENGINE_PORT="${DOC_HAUS_ENGINE_PORT:-4097}"
DOC_HAUS_INGEST_PORT="${DOC_HAUS_INGEST_PORT:-4501}"
DOC_HAUS_WEB_PORT="${DOC_HAUS_WEB_PORT:-5174}"
if [ -L "$app" ] || { [ -e "$app" ] && [ ! -d "$app" ]; }; then
  echo "doc.haus: the application path is not a regular app folder." >&2
  exit 1
fi
mkdir -p "$(dirname "$app")"
stage=$(mktemp -d "$(dirname "$app")/.doc.haus-build.XXXXXX")
trap 'rm -rf "$stage"' EXIT
candidate="$stage/doc.haus.app"
bun=$(command -v "${BUN_BIN:-bun}" 2>/dev/null || true)
if [ -n "$bun" ]; then bun="$(cd "$(dirname "$bun")" && pwd)/$(basename "$bun")"; fi
word_helper=
if [ -f "$source_scripts/word-open.ts" ] && [ -n "$bun" ]; then
  if (cd "$source_scripts/.." && "$bun" build "$source_scripts/word-open.ts" --target bun --outfile "$stage/word-open.js") >"$stage/helper-build.log" 2>&1; then
    word_helper="$stage/word-open.js"
  elif [ "${DOCHAUS_DESKTOP_HELPER:-}" = 1 ]; then
    echo "doc.haus: the desktop Word helper could not be built. Use Settings > Add-ins to try again." >&2
    exit 1
  fi
fi
if [ "${DOCHAUS_DESKTOP_HELPER:-}" = 1 ] && [ -z "$word_helper" ]; then
  echo "doc.haus: the desktop Word helper is unavailable. Use Settings > Add-ins to try again." >&2
  exit 1
fi
if [ -z "$word_helper" ] && [ -f "$app/Contents/Resources/word-open.js" ]; then
  echo "doc.haus: keeping the existing desktop Word helper; startup will retry its update."
  exit 0
fi
if [ -n "${DOCHAUS_DESKTOP_METADATA:-}" ]; then printf '%s\n' "$DOCHAUS_DESKTOP_METADATA" > "$stage/installation.json"; fi
{
  printf '#!/bin/bash\nexport DOCHAUS_INSTALL_DIR=%q\n' "$repository"
  if [ -n "$word_helper" ]; then printf 'export DOCHAUS_BUN_BIN=%q\n' "$bun"; fi
  # Only installation paths and ports survive Finder's separate environment.
  # Never copy provider credentials or the caller's general environment.
  for variable in DOCHAUS_DATA_ROOT WORKSPACE_ROOT DOCHAUS_RUNTIME_DIR DOC_HAUS_ENGINE_PORT DOC_HAUS_INGEST_PORT DOC_HAUS_WEB_PORT; do
    if [ -n "${!variable:-}" ]; then printf 'export %s=%q\n' "$variable" "${!variable}"; fi
  done
} > "$stage/installation.sh"
icon="$source_scripts/icons/doc-haus.icns"
if [ ! -f "$icon" ]; then icon="$repository/scripts/icons/doc-haus.icns"; fi
# Compare inputs before compiling/signing: an ordinary startup must not churn
# the app's identity, Automation consent or LaunchServices caches.
stamp=$({
  cat "$source_scripts/macos-app.sh" "$source_scripts/macos-app.applescript" "$source_scripts/macos-app-launch.sh" "$stage/installation.sh"
  if [ -n "$word_helper" ]; then cat "$word_helper"; fi
  if [ -f "$stage/installation.json" ]; then cat "$stage/installation.json"; fi
  if [ -f "$icon" ]; then cat "$icon"; fi
} | shasum -a 256 | cut -d ' ' -f 1)
if [ "$stamp" = "$(cat "$app/Contents/Resources/launcher-inputs.sha256" 2>/dev/null || true)" ] &&
   codesign --verify --deep --strict "$app" >/dev/null 2>&1; then
  echo "doc.haus: Mac app is already current."
  exit 0
fi
osacompile -s -o "$candidate" "$source_scripts/macos-app.applescript"
resources="$candidate/Contents/Resources"
mkdir -p "$resources"
cp "$stage/installation.sh" "$resources/installation.sh"
if [ -f "$stage/installation.json" ]; then cp "$stage/installation.json" "$resources/installation.json"; fi
printf '%s\n' "$stamp" > "$resources/launcher-inputs.sha256"
cp "$source_scripts/macos-app-launch.sh" "$resources/open-doc-haus.sh"
if [ -n "$word_helper" ]; then
  cp "$word_helper" "$resources/word-open.js"
  cat > "$resources/open-word.sh" <<'WORD'
#!/bin/bash
set -euo pipefail
resources="$(cd "$(dirname "$0")" && pwd)"
. "$resources/installation.sh"
cd "$DOCHAUS_INSTALL_DIR"
. scripts/data-paths.sh
configure_data_paths "$PWD"
exec "$DOCHAUS_BUN_BIN" "$resources/word-open.js" --runtime "$DOCHAUS_RUNTIME_DIR" "$@"
WORD
  chmod +x "$resources/open-word.sh"
fi
cat > "$resources/doc.haus.command" <<'COMMAND'
#!/bin/bash
set -euo pipefail
. "$(dirname "$0")/installation.sh"
exec "$DOCHAUS_INSTALL_DIR/start.sh"
COMMAND
chmod +x "$resources/doc.haus.command" "$resources/open-doc-haus.sh"
if [ -f "$icon" ]; then cp "$icon" "$resources/AppIcon.icns"; fi
plist="$candidate/Contents/Info.plist"
plutil -replace CFBundleIdentifier -string haus.doc.start "$plist"
plutil -replace CFBundleName -string doc.haus "$plist"
plutil -replace CFBundleDisplayName -string doc.haus "$plist"
plutil -replace CFBundleIconFile -string AppIcon "$plist"
plutil -replace CFBundleURLTypes -json '[{"CFBundleURLName":"haus.doc.launch","CFBundleURLSchemes":["doc.haus"]}]' "$plist"
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
  recovery="$(dirname "$app")/.doc.haus-recovery"
  if [ -L "$recovery" ] || { [ -e "$recovery" ] && [ ! -d "$recovery" ]; }; then
    echo "doc.haus: the app recovery path is not a regular folder." >&2
    exit 1
  fi
  mkdir -p -m 700 "$recovery"
  backup="$recovery/$(basename "$app").previous-$(date +%Y%m%d%H%M%S)-$$"
  mv "$app" "$backup"
fi
if ! mv "$candidate" "$app"; then
  if [ -n "$backup" ]; then mv "$backup" "$app"; fi
  exit 1
fi
touch "$app"
echo "doc.haus: Mac app ready at $app"
if [ -n "$backup" ]; then echo "doc.haus: previous app kept at $backup"; fi
