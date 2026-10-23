# Setup

Use Git. Run ./start.sh from a clone. The launcher uses Bun 1.4.0 or newer from
PATH, then the official home install (~/.bun/bin/bun);
set BUN_BIN=/absolute/path/to/bun to choose another executable. Without a usable
Bun it offers to run the official installer from https://bun.sh. An explicit
BUN_BIN is never replaced. The package and
container pins provide reproducible baselines and do not restrict local startup.
The launcher installs the root workspace and the separate ingest dependency graph with frozen lockfiles.
It never needs provider credentials just to boot. Ports default to engine 4097,
ingest 4501 and web 5174. Use DOC_HAUS_ENGINE_PORT, DOC_HAUS_INGEST_PORT and
DOC_HAUS_WEB_PORT to run an isolated copy, together with a separate data root.

```sh
DOCHAUS_DATA_ROOT=/absolute/path/to/dochaus-data ./start.sh
```

First-run setup assumes this is your first installation:

1. Choose your model provider from the dropdown. Detected options appear first;
   choose another provider or a local/custom endpoint if needed.
2. Connect using the provider's native API key, OAuth, command or environment
   method. Available project, profile, region and location settings are prefilled.
   Detected environment credentials need no copying or re-entry.
3. Confirm a primary and fast model from that provider; the same model is allowed.
   Research and firm details are optional and can be skipped.

The last setup step offers a sample matter to explore: a fictional engagement
letter, ready for cited questions and a legal review. It is pre-selected only
when the workspace has no matters and is added when you finish setup. Add it
later from Settings → Matter defaults; the sample is labelled in the matters list
and is deleted like any matter. Scripts and test harnesses can add it with
`POST /sample-matter` on the ingest service (default port 4501). The request is
idempotent: it returns the existing sample rather than adding a second one.

Detection uses OpenCode's native provider/integration/model services and works
without another OpenCode installation. The launcher inherits provider environment
variables. Native local Ollama/LM Studio discovery is reused; a custom
OpenAI-compatible endpoint is a fallback.

Bedrock uses native AWS profile/environment hints and the AWS SDK credential
chain. Shared-profile files can be shown as a source; choose a profile when no
active profile is configured. Vertex uses ADC plus a configured project or
supported project environment variable. ADC alone does not select a project.
Missing cloud settings can be entered in the connection stage. Saving them does
not verify model access. Host providers require an explicit successful model test
before their models become selectable. Tests may consume request allowance;
opening setup does not silently test a model. Detection does not establish quota,
organization-policy approval or subscription access.

Set up later leaves required setup unfinished. Use Settings to resume or change
provider options. OpenCode Go and OpenCode Zen are separate native integrations;
Console login does not establish a Go subscription. If you already use OpenCode,
setup checks for matching saved sign-ins after you select a provider. Only
matching connections are offered, and copying requires your explicit selection.
The source is preserved; global configuration, plugins, instructions and sessions
are never imported. An existing sign-in remains subject to its account policy.

After setup is complete and the launch notice is accepted, opening the app checks
access to the saved primary and fast models with a small fixed test request.
Identical models share one check; recent checks are briefly reused across tabs.
No matter content is sent and no drafting is resumed. These requests may consume
provider allowance. Saved connections remain listed in Settings, with their
latest verification result. A rejected sign-in shows a persistent reconnect
prompt; an expired API key requires a replacement key from the provider console.
Network or account-allowance failures ask you to check the connection instead.
Reconnect or save new model choices to check again, or use Check again in the
prompt. Authentication failures during assistant work also show the prompt.
After reconnecting, return to the draft and choose Continue when ready.

New installations default to Entire web; existing saved research choices are
preserved. Approved sources only shows the built-in host/domain list and lets you
add trusted hosts. Optional keys start empty and CourtListener is off. Enable the
connector to enter its token. Direct official retrieval and keyless search remain
available. Research tests run only when requested; Exa's minimal test uses its
search allowance. Tokens stay out of model instructions.
Only the shipped local research server and explicitly consented CourtListener
endpoint are managed. A failed consent filter prevents managed readiness;
plugin/native failures must be diagnosed, not treated as an unconditional boundary.

## Host tools

After the ownership and port checks, the launcher looks for the optional tools
the stack calls on PATH and offers to install any that are missing or too old:

- Deno 2+ for the shipped Python MCP tool. Its first activation downloads the
  declared Deno/Pyodide dependencies.
- Python 3.10+ and markitdown (or uvx) for the PDF text fallback; markitdown
  installs through pipx.
- Poppler and Tesseract for OCR of scanned PDFs.
- LibreOffice Writer, with fontconfig and the Carlito font, for PDF to DOCX
  conversion and Calibri-compatible rendering.
- lsof and pgrep (Linux) for the launcher's port checks and shutdown.

Each install is shown before it runs and Enter accepts. macOS uses Homebrew and
installs Homebrew first when absent; Linux uses apt-get, dnf or pacman with sudo.
Deno installs with its official installer where the distribution has no package.
The launcher adds ~/.deno/bin, ~/.local/bin and the macOS LibreOffice app to its
own PATH so tools installed there work without a new shell.

A declined set is recorded in the runtime directory (declined-tools) and not
offered again until another tool goes missing. Set DOCHAUS_INSTALL_TOOLS=ask to
be asked again, yes to install without asking (also without a terminal) or no to
never install. Without a terminal the launcher only reports missing tools.
SKIP_INSTALL=1 skips these checks along with dependency installation.
Use docker compose up --build for the bundled local Linux toolchain. Only host
127.0.0.1:5173 is published; /engine and /ingest proxy to internal loopback
services and preserve native /api paths and streaming. /wasm hosts Docxodus assets.

The launcher identifies an existing runtime owner. Preserve an existing stack
unless its owner has explicitly agreed to a restart. Never share one runtime
between simultaneous engines.
