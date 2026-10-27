#!/usr/bin/env bash
# doc.haus installer for macOS and Linux:
#   curl -fsSL https://doc.haus/install.sh | bash
# Installs Git when it is missing, downloads doc.haus to ~/doc-haus (or
# DOCHAUS_DIR) and starts it. Running it again starts the existing copy.
set -euo pipefail
repository=https://github.com/sure-scale/doc-haus.git
dir="${DOCHAUS_DIR:-$HOME/doc-haus}"

as_root() { if [ "$(id -u)" = 0 ]; then "$@"; else sudo "$@"; fi; }

if ! git --version >/dev/null 2>&1; then
  if [ "$(uname -s)" = Darwin ]; then
    # Git comes with Apple's Command Line Tools, which install through a window.
    # Wait for them while that window is open, then carry on.
    xcode-select --install >/dev/null 2>&1 || true
    echo "doc.haus: Git comes with Apple's Command Line Tools. Click Install in the window that opened; doc.haus carries on once they are installed." >&2
    waited=0
    until xcode-select -p >/dev/null 2>&1 && git --version >/dev/null 2>&1; do
      if [ "$waited" -ge 15 ] && ! pgrep -f "Install Command Line Developer Tools" >/dev/null; then
        echo "doc.haus: the Command Line Tools were not installed. Run this command again to retry." >&2
        exit 1
      fi
      sleep 5; waited=$((waited + 5))
    done
  else
    echo "doc.haus: installing Git" >&2
    if command -v apt-get >/dev/null; then as_root apt-get update && as_root apt-get install -y git
    elif command -v dnf >/dev/null; then as_root dnf install -y git
    elif command -v pacman >/dev/null; then as_root pacman -S --noconfirm git
    else echo "doc.haus needs Git. Install it with your package manager, then run this command again." >&2; exit 1
    fi
  fi
fi

if [ -e "$dir" ]; then
  case "$(git -C "$dir" remote get-url origin 2>/dev/null)" in
    *sure-scale/doc-haus*) echo "doc.haus: starting the copy in $dir" >&2 ;;
    *) echo "doc.haus: $dir already exists and is not a copy of doc.haus. Set DOCHAUS_DIR to another folder, then run this command again." >&2; exit 1 ;;
  esac
else
  git clone "$repository" "$dir"
fi

dir="$(cd "$dir" && pwd)"
# A doc.haus entry to start it again later, which opens a terminal window
# running start.sh: an app in ~/Applications on macOS, which Spotlight and
# Launchpad find, and an applications-menu entry on Linux. Both show the logo
# from scripts/icons, which copies made before it was added do not have.
if [ "$(uname -s)" = Darwin ]; then
  app="$HOME/Applications/doc.haus.app"
  mkdir -p "$app/Contents/MacOS" "$app/Contents/Resources"
  {
    printf '#!/bin/bash\n'
    # Finder does not inherit the install terminal's environment. Keep only the
    # selected data locations and ports, so reopening uses the same installation.
    # Never copy provider credentials or the caller's general environment here.
    for variable in DOCHAUS_DATA_ROOT WORKSPACE_ROOT DOCHAUS_RUNTIME_DIR DOC_HAUS_ENGINE_PORT DOC_HAUS_INGEST_PORT DOC_HAUS_WEB_PORT; do
      if [ -n "${!variable:-}" ]; then printf 'export %s=%q\n' "$variable" "${!variable}"; fi
    done
    printf 'exec %q\n' "$dir/start.sh"
  } > "$app/Contents/Resources/doc.haus.command"
  # shellcheck disable=SC2016 # $0 belongs to the app's own script
  printf '#!/bin/bash\nexec open -a Terminal "$(dirname "$0")/../Resources/doc.haus.command"\n' > "$app/Contents/MacOS/doc.haus"
  chmod +x "$app/Contents/Resources/doc.haus.command" "$app/Contents/MacOS/doc.haus"
  if [ -f "$dir/scripts/icons/doc-haus.icns" ]; then cp "$dir/scripts/icons/doc-haus.icns" "$app/Contents/Resources/AppIcon.icns"; fi
  cat > "$app/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>CFBundleExecutable</key><string>doc.haus</string>
  <key>CFBundleIconFile</key><string>AppIcon</string>
  <key>CFBundleIdentifier</key><string>haus.doc.start</string>
  <key>CFBundleName</key><string>doc.haus</string>
  <key>CFBundlePackageType</key><string>APPL</string>
</dict></plist>
PLIST
  # Finder shows a changed icon once the app itself has changed.
  touch "$app"
else
  applications="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
  mkdir -p "$applications"
  printf '[Desktop Entry]\nType=Application\nName=doc.haus\nComment=Start doc.haus\nExec="%s/start.sh"\nPath=%s\nIcon=%s/scripts/icons/doc-haus.png\nTerminal=true\nCategories=Office;\n' "$dir" "$dir" "$dir" > "$applications/doc-haus.desktop"
fi

cd "$dir"
export DOCHAUS_SETUP_WORD_ADDIN=1
# Under curl | bash, standard input is the script, so give the launcher the
# terminal for its install questions.
if { : </dev/tty; } 2>/dev/null; then exec ./start.sh "$@" </dev/tty; fi
exec ./start.sh "$@"
