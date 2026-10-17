# Sourced by the launcher. Selects Bun and checks the host tools doc.haus calls
# on PATH, offering to install whatever is missing or too old. Every install is
# asked first and Enter means yes. Without a terminal nothing is installed
# unless DOCHAUS_INSTALL_TOOLS=yes; DOCHAUS_INSTALL_TOOLS=no never installs and
# DOCHAUS_INSTALL_TOOLS=ask asks again about tools declined before.

# Ask a yes/no question that defaults to yes.
confirm_install() {
  local answer
  case "${DOCHAUS_INSTALL_TOOLS:-}" in
    yes) return 0 ;;
    no) return 1 ;;
  esac
  [ -t 0 ] || return 1
  printf '%s [Y/n] ' "$1" >&2
  read -r answer || answer=n
  case "$answer" in
    ""|y|Y|yes|YES|Yes) return 0 ;;
  esac
  return 1
}

# Print each install command before running it, so the user sees what changes.
run_install() {
  echo "+ $*" >&2
  "$@"
}

as_root() {
  if [ "$(id -u)" = 0 ]; then
    run_install "$@"
  elif command -v sudo >/dev/null 2>&1; then
    run_install sudo "$@"
  else
    echo "doc.haus: installing $* needs root, and sudo is not available." >&2
    return 1
  fi
}

package_manager() {
  local manager
  if [ "$(uname -s)" = Darwin ]; then
    echo brew
    return 0
  fi
  for manager in apt-get dnf pacman; do
    if command -v "$manager" >/dev/null 2>&1; then
      echo "$manager"
      return 0
    fi
  done
}

# Install system packages ($2...) with package manager $1.
install_packages() {
  local manager=$1
  shift
  [ "$#" -gt 0 ] || return 0
  case "$manager" in
    brew) run_install brew install "$@" ;;
    apt-get) as_root apt-get update && as_root env DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends "$@" ;;
    dnf) as_root dnf install -y "$@" ;;
    pacman) as_root pacman -S --needed --noconfirm "$@" ;;
    *) return 1 ;;
  esac
}

# The official Bun and Deno installers need curl and unzip.
ensure_download_tools() {
  local missing=() tool
  for tool in curl unzip; do
    command -v "$tool" >/dev/null 2>&1 || missing+=("$tool")
  done
  [ "${#missing[@]}" -gt 0 ] || return 0
  install_packages "$(package_manager)" "${missing[@]}"
}

BUN_VERSION_PATTERN='^([0-9]+)\.([0-9]+)\.[0-9]+(-[a-zA-Z0-9.-]+)?(\+[a-zA-Z0-9.-]+)?$'

bun_supported() {
  [[ "$1" =~ $BUN_VERSION_PATTERN ]] || return 1
  [ "${BASH_REMATCH[1]}" -gt 1 ] || { [ "${BASH_REMATCH[1]}" -eq 1 ] && [ "${BASH_REMATCH[2]}" -ge 4 ]; }
}

# Set BUN_BIN to the absolute path of $1 and BUN_FOUND to its version, or
# leave BUN_FOUND empty when it cannot run.
select_bun() {
  BUN_BIN=$1
  BUN_FOUND=
  if [ -n "$BUN_BIN" ] && [[ "$BUN_BIN" != */* ]]; then
    BUN_BIN="$(command -v "$BUN_BIN" || true)"
  fi
  if [ -n "$BUN_BIN" ] && [ -x "$BUN_BIN" ] && [ ! -d "$BUN_BIN" ]; then
    if [[ "$BUN_BIN" != /* ]]; then
      BUN_BIN="$(cd "$(dirname "$BUN_BIN")" && pwd -P)/$(basename "$BUN_BIN")"
    fi
    BUN_FOUND="$("$BUN_BIN" --version 2>/dev/null)" || BUN_FOUND=
  fi
}

# Choose Bun 1.4.0+ as BUN_BIN: an explicit BUN_BIN, else PATH, else the
# official home install. Prefer the installed runtime; build/CI pins remain a
# reproducible baseline, not a reason to reject a working newer Bun. Without a
# usable Bun, offer the official installer; an explicit BUN_BIN is never replaced.
ensure_bun() {
  local explicit=${BUN_BIN:-} home="${BUN_INSTALL:-$HOME/.bun}/bin/bun" path_bin path_found runs=
  if [ -n "$explicit" ]; then
    select_bun "$explicit"
  else
    select_bun "$(command -v bun || true)"
    if ! bun_supported "$BUN_FOUND"; then
      path_bin=$BUN_BIN path_found=$BUN_FOUND
      select_bun "$home"
      if ! bun_supported "$BUN_FOUND" && [ -n "$path_found" ]; then
        BUN_BIN=$path_bin BUN_FOUND=$path_found
      fi
    fi
  fi
  bun_supported "$BUN_FOUND" && return 0
  if [[ "$BUN_FOUND" =~ $BUN_VERSION_PATTERN ]]; then
    runs=1
    echo "doc.haus requires Bun 1.4.0 or newer (found: $BUN_FOUND)." >&2
  else
    echo "doc.haus could not run Bun (${BUN_BIN:-not found on PATH})." >&2
  fi
  if [ -z "$explicit" ] && confirm_install "doc.haus: install the latest Bun from https://bun.sh now?"; then
    if ensure_download_tools && run_install bash -c 'set -o pipefail; curl -fsSL https://bun.sh/install | bash'; then
      select_bun "$home"
      bun_supported "$BUN_FOUND" && return 0
    fi
    echo "doc.haus: the Bun install did not produce a working Bun 1.4.0+ at $home." >&2
  fi
  if [ -z "$runs" ]; then
    echo "Install Bun or set BUN_BIN to a working Bun executable, then run ./start.sh again." >&2
  else
    echo "Update Bun or set BUN_BIN to a Bun 1.4.0+ executable, then run ./start.sh again." >&2
  fi
  return 1
}

# Installers put some tools outside the PATH of a shell opened before they ran.
# Append those locations so the engine and ingest find the tools straight away;
# appending keeps every tool already on PATH first.
add_tool_paths() {
  local dir brew
  for brew in /opt/homebrew/bin/brew /usr/local/bin/brew /home/linuxbrew/.linuxbrew/bin/brew; do
    if [ -x "$brew" ] && ! command -v brew >/dev/null 2>&1; then
      PATH="$PATH:$(dirname "$brew")"
    fi
  done
  for dir in "${DENO_INSTALL:-$HOME/.deno}/bin" "$HOME/.local/bin" /Applications/LibreOffice.app/Contents/MacOS; do
    case ":$PATH:" in
      *":$dir:"*) ;;
      *)
        if [ -d "$dir" ] && { [ "$dir" != /Applications/LibreOffice.app/Contents/MacOS ] || ! command -v soffice >/dev/null 2>&1; }; then
          PATH="$PATH:$dir"
        fi
        ;;
    esac
  done
  export PATH
}

TOOL_IDS="deno python markitdown poppler tesseract libreoffice fonts processes"

tool_label() {
  case "$1" in
    deno) echo "Deno 2+ (runs the shipped Python tool)" ;;
    python) echo "Python 3.10+ (runs markitdown)" ;;
    markitdown) echo "markitdown (PDF text fallback)" ;;
    poppler) echo "Poppler (renders PDF pages for OCR)" ;;
    tesseract) echo "Tesseract (OCR for scanned PDFs)" ;;
    libreoffice) echo "LibreOffice Writer (PDF to DOCX conversion)" ;;
    fonts) echo "Carlito font (Calibri-compatible DOCX rendering)" ;;
    processes) echo "lsof and pgrep (launcher port checks and shutdown)" ;;
  esac
}

# macOS ships /usr/bin/python3 as a stub that opens an installer dialog until
# the Command Line Tools exist, so never run it there.
python_usable() {
  local python
  python="$(command -v python3 || true)"
  [ -n "$python" ] || return 1
  [ "$(uname -s)" != Darwin ] || [ "$python" != /usr/bin/python3 ] || xcode-select -p >/dev/null 2>&1
}

# Print why tool $1 needs installing, or nothing when it is usable.
tool_problem() {
  local version major minor fonts
  case "$1" in
    deno)
      if ! command -v deno >/dev/null 2>&1; then echo "not found"; return 0; fi
      version="$(deno --version 2>/dev/null || true)"
      version=${version#deno }
      version=${version%%[!0-9.]*}
      if ! [ "${version%%.*}" -ge 2 ] 2>/dev/null; then echo "found ${version:-an unknown version}, needs 2 or newer"; fi
      ;;
    python)
      if ! python_usable; then echo "not found"; return 0; fi
      version="$(python3 -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null || true)"
      major=${version%%.*} minor=${version#*.}
      if ! { [ "$major" -gt 3 ] || { [ "$major" -eq 3 ] && [ "$minor" -ge 10 ]; }; } 2>/dev/null; then
        echo "found ${version:-an unknown version}, needs 3.10 or newer"
      fi
      ;;
    markitdown)
      # Ingest runs markitdown through uvx when it is not installed directly.
      command -v markitdown >/dev/null 2>&1 || command -v uvx >/dev/null 2>&1 || echo "not found"
      ;;
    poppler) command -v pdftoppm >/dev/null 2>&1 || echo "not found" ;;
    tesseract) command -v tesseract >/dev/null 2>&1 || echo "not found" ;;
    libreoffice) command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1 || echo "not found" ;;
    fonts)
      fonts="$(fc-list : family 2>/dev/null || true)"
      if [ "$(uname -s)" = Darwin ]; then
        fonts="$fonts $(ls "$HOME/Library/Fonts" /Library/Fonts 2>/dev/null || true)"
      fi
      case "$fonts" in
        *[Cc]arlito*) ;;
        *) echo "not found" ;;
      esac
      ;;
    processes)
      command -v lsof >/dev/null 2>&1 && command -v pgrep >/dev/null 2>&1 || echo "not found"
      ;;
  esac
}

# Print what installs tool $1 with package manager $2: system packages, plus
# cask:<name> (Homebrew cask), pipx:<name>, deno-installer, deno-upgrade and
# brew-upgrade (Deno installed by Homebrew).
tool_install_plan() {
  local id=$1 manager=$2
  if [ "$id" = deno ] && command -v deno >/dev/null 2>&1; then
    if [ "$manager" = brew ] && brew list deno >/dev/null 2>&1; then echo brew-upgrade; else echo deno-upgrade; fi
    return 0
  fi
  case "$manager:$id" in
    brew:deno) echo deno ;;
    brew:python) echo python ;;
    brew:markitdown) echo "pipx pipx:markitdown[pdf]" ;;
    brew:poppler) echo poppler ;;
    brew:tesseract) echo tesseract ;;
    brew:libreoffice) echo cask:libreoffice ;;
    brew:fonts) echo cask:font-carlito ;;
    apt-get:deno|dnf:deno) echo "curl unzip deno-installer" ;;
    apt-get:python) echo "python3 python3-venv" ;;
    apt-get:markitdown|dnf:markitdown) echo "pipx pipx:markitdown[pdf]" ;;
    apt-get:poppler|dnf:poppler) echo poppler-utils ;;
    apt-get:tesseract) echo tesseract-ocr ;;
    apt-get:libreoffice|dnf:libreoffice) echo libreoffice-writer ;;
    apt-get:fonts) echo "fontconfig fonts-crosextra-carlito" ;;
    apt-get:processes) echo "lsof procps" ;;
    dnf:python) echo python3 ;;
    dnf:tesseract) echo "tesseract tesseract-langpack-eng" ;;
    dnf:fonts) echo "fontconfig google-carlito-fonts" ;;
    dnf:processes|pacman:processes) echo "lsof procps-ng" ;;
    pacman:deno) echo deno ;;
    pacman:python) echo python ;;
    pacman:markitdown) echo "python-pipx pipx:markitdown[pdf]" ;;
    pacman:poppler) echo poppler ;;
    pacman:tesseract) echo "tesseract tesseract-data-eng" ;;
    pacman:libreoffice) echo libreoffice-fresh ;;
    pacman:fonts) echo "fontconfig ttf-carlito" ;;
  esac
}

# Print the ids of tools that are missing or too old.
missing_tools() {
  local id list=
  for id in $TOOL_IDS; do
    if [ -n "$(tool_problem "$id")" ]; then list="${list:+$list }$id"; fi
  done
  echo "$list"
}

# Install tools $2... with package manager $1. Failures are reported, never fatal:
# these tools are optional and the stack starts without them.
install_tools() {
  local manager=$1 id step packages=() casks=() pipx=() deno_install='' deno_upgrade=''
  shift
  for id in "$@"; do
    for step in $(tool_install_plan "$id" "$manager"); do
      case "$step" in
        cask:*) casks+=("${step#cask:}") ;;
        pipx:*) pipx+=("${step#pipx:}") ;;
        deno-installer) deno_install=1 ;;
        deno-upgrade) deno_upgrade=deno ;;
        brew-upgrade) deno_upgrade=brew ;;
        *) packages+=("$step") ;;
      esac
    done
  done
  if [ "$manager" = brew ] && ! command -v brew >/dev/null 2>&1; then
    echo "doc.haus: installing Homebrew (https://brew.sh) first; it may ask for your password." >&2
    run_install /bin/bash -c 'script="$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" && /bin/bash -c "$script"' || true
    add_tool_paths
    if ! command -v brew >/dev/null 2>&1; then
      echo "doc.haus: Homebrew is not available, so the tools were not installed." >&2
      return 0
    fi
  fi
  install_packages "$manager" ${packages[@]+"${packages[@]}"} || echo "doc.haus: some packages did not install." >&2
  if [ "${#casks[@]}" -gt 0 ]; then
    run_install brew install --cask "${casks[@]}" || echo "doc.haus: some applications did not install." >&2
  fi
  if [ -n "$deno_install" ]; then
    run_install bash -c 'set -o pipefail; curl -fsSL https://deno.land/install.sh | sh -s -- -y' || echo "doc.haus: Deno did not install." >&2
  fi
  case "$deno_upgrade" in
    brew) run_install brew upgrade deno || echo "doc.haus: Deno did not update." >&2 ;;
    deno) run_install deno upgrade || echo "doc.haus: Deno did not update." >&2 ;;
  esac
  add_tool_paths
  for id in ${pipx[@]+"${pipx[@]}"}; do
    if command -v pipx >/dev/null 2>&1; then
      run_install pipx install "$id" || echo "doc.haus: $id did not install." >&2
    else
      echo "doc.haus: pipx is not available, so $id was not installed." >&2
    fi
  done
  add_tool_paths
}

# Offer to install missing optional tools. A declined set is remembered in
# $1/declined-tools and not asked about again until another tool goes missing.
ensure_host_tools() {
  local declined_file="$1/declined-tools" missing id manager question declined='' new=''
  missing="$(missing_tools)"
  if [ -z "$missing" ]; then
    rm -f "$declined_file"
    return 0
  fi
  [ ! -f "$declined_file" ] || declined=" $(cat "$declined_file") "
  for id in $missing; do
    case "$declined" in *" $id "*) ;; *) new=1 ;; esac
  done
  if [ -z "$new" ] && [ -z "${DOCHAUS_INSTALL_TOOLS:-}" ]; then
    echo "doc.haus: optional tools not installed: $missing (set DOCHAUS_INSTALL_TOOLS=ask to be asked again)" >&2
    return 0
  fi
  echo "doc.haus: some optional tools are missing or out of date:" >&2
  for id in $missing; do
    echo "  - $(tool_label "$id"): $(tool_problem "$id")" >&2
  done
  manager="$(package_manager)"
  if [ -z "$manager" ]; then
    echo "doc.haus: no supported package manager (Homebrew, apt-get, dnf or pacman) was found; install these tools manually. The stack starts without them." >&2
    return 0
  fi
  if [ "${DOCHAUS_INSTALL_TOOLS:-}" != yes ] && { [ "${DOCHAUS_INSTALL_TOOLS:-}" = no ] || [ ! -t 0 ]; }; then
    echo "doc.haus: starting without them. Run ./start.sh in a terminal to install them, or set DOCHAUS_INSTALL_TOOLS=yes." >&2
    return 0
  fi
  question="doc.haus: install them with $manager now?"
  if [ "$manager" = brew ] && ! command -v brew >/dev/null 2>&1; then
    question="doc.haus: install Homebrew (https://brew.sh), then these tools, now?"
  fi
  if ! confirm_install "$question"; then
    echo "$missing" > "$declined_file"
    echo "doc.haus: starting without them; you will not be asked again for these tools (set DOCHAUS_INSTALL_TOOLS=ask to change that)." >&2
    return 0
  fi
  rm -f "$declined_file"
  # shellcheck disable=SC2086 # ids are single words
  install_tools "$manager" $missing
  missing="$(missing_tools)"
  if [ -n "$missing" ]; then
    echo "doc.haus: still missing: $missing; the stack starts without them." >&2
  fi
  return 0
}
