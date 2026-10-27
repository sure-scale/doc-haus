# doc.haus

A local legal-document workspace built on [OpenCode](https://github.com/anomalyco/opencode).
Upload documents into matters, ask source-grounded questions, and prepare drafts
and redlines for lawyer review. Generated work requires professional review.

## Start locally

One command installs Git if it is missing, downloads doc.haus to `~/doc-haus`
(or `DOCHAUS_DIR`) and starts it. Run it again to start doc.haus later.

```sh
curl -fsSL https://doc.haus/install.sh | bash
```

On Windows, run this in PowerShell:

```powershell
irm https://doc.haus/install.ps1 | iex
```

These run [scripts/install.sh](scripts/install.sh) and
[scripts/install.ps1](scripts/install.ps1). To install by hand, install Git,
clone this repository, then run:

```sh
git clone https://github.com/sure-scale/doc-haus.git
cd doc-haus
./start.sh
```

On macOS, the installer adds **doc.haus** to Applications. Clicking it opens
the running installation or starts it in Terminal when needed, then the launcher
exits. Chrome-family browsers and Safari can reuse an existing doc.haus tab
without leaving the matter you have open; macOS asks for Automation permission
first. Other browsers, or denied access, use normal URL opening. Terminal keeps
the local services running; if you stop them, click doc.haus to start again.
The app keeps data-directory and engine, ingest or web port overrides selected
during installation. Separate copies have separate settings and matters; the app
does not import another copy's setup. After updating an existing installation,
run the installer again to refresh an older Mac shortcut.

The launcher checks for Bun 1.4.0 or newer and the document tools (Deno 2,
Python 3.10+, markitdown, Poppler, Tesseract, LibreOffice Writer and the Carlito
font). It lists anything missing or out of date and offers to install it; press
Enter to accept. macOS uses Homebrew (offered first if absent); Linux uses
apt-get, dnf or pacman. The launcher then installs frozen dependencies and opens
the web app at http://localhost:5174. It runs engine, ingest and web on separate
local ports.
A new installation needs no credentials to boot. Accept the launch notice,
choose a provider (detected options appear first), connect with its offered sign-in
method and detected defaults, then confirm primary/fast models in setup. Research and firm
or jurisdiction details are optional. The last step can add a fictional sample
matter to explore. Opening setup does not make a paid model
request; connection tests run only when requested. Provider access and quota
remain subject to the provider's account and organization policies.

Only Bun is required, and the launcher installs it. The stack starts without
declined tools, and a declined set is not offered again until another tool goes
missing. The container includes every tool. See [setup](docs/setup.md).

## Start on Windows

On Windows 10 or 11, install Git for Windows, clone the repository, then run
`start.cmd` from its folder in Command Prompt or PowerShell (`.\start.cmd`). It
finds x64 Bun 1.4.0 or newer (Windows on Arm runs it emulated), offers to
install it when missing, and then starts the stack like `start.sh`. In Git Bash,
`./start.sh` hands over to `start.cmd`. Press Ctrl+C to stop it. See
[setup](docs/setup.md#start-on-windows).

## Edit in Word

Use **Open in Word** on a document, then **Save to doc.haus** in the Word pane
and wait for confirmation before closing. The pane checks versions and claims
editing on the first save; doc.haus blocks competing writes until you finish or
release the lock. Background sync is best effort and
can combine saves. The installers offer add-in setup; existing users and firms
can follow [Word add-in setup](docs/word-addin.md).

## Local container

```sh
docker compose up --build
```

Open http://localhost:5173. Only web is published, on host loopback. Engine and
ingest stay on container loopback behind /engine and /ingest proxies. Named
workspace and runtime volumes retain documents, native credentials, sessions
and settings across recreation. This is a local self-host configuration.
Internet-facing or team hosting needs a separate deployment design.

## Settings, data and updates

Saving settings leaves the shipped dochaus/opencode.json unchanged. User model
and provider options and agent/skill toggles live in runtime configuration;
credentials use native storage. Custom builder definitions have separate backup
requirements described in the storage guide.
See [storage and backup](docs/storage.md) and [upgrades](docs/updates.md), including
the explicit migration path for older installations with tracked settings.

Detailed provider/tool captures are **off by default**. Troubleshooting capture
requires DOCHAUS_PROVIDER_CAPTURE_DIR set to a private absolute directory;
DOCHAUS_PROVIDER_CAPTURE_DIR=off disables it without deleting prior captures.
Captures can contain confidential document and model content. Ordinary native
sessions and sanitized failure diagnostics remain durable independently.

[Contributing](CONTRIBUTING.md) · [Security](SECURITY.md) ·
[Licensing](LICENSING.md) · [Attribution](ATTRIBUTION.md)

The root `install` script belongs to upstream OpenCode. Use `start.sh` for doc.haus.
