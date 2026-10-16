# Setup

Use Git and Bun 1.4.0 or newer. Run ./start.sh from a clone. The launcher uses
Bun from PATH;
set BUN_BIN=/absolute/path/to/bun to choose another executable. The package and
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

New installations default to Entire web; existing saved research choices are
preserved. Approved sources only shows the built-in host/domain list and lets you
add trusted hosts. Optional keys start empty and CourtListener is off. Enable the
connector to enter its token. Direct official retrieval and keyless search remain
available. Research tests run only when requested; Exa's minimal test uses its
search allowance. Tokens stay out of model instructions.
Only the shipped local research server and explicitly consented CourtListener
endpoint are managed. A failed consent filter prevents managed readiness;
plugin/native failures must be diagnosed, not treated as an unconditional boundary.

The shipped Python MCP tool requires Deno 2 (the container includes it).
Its first activation downloads the declared Deno/Pyodide dependencies.
For PDF/OCR fallback install Poppler, Tesseract, Python 3 and markitdown.
For DOCX rendering install LibreOffice Writer, fontconfig and Carlito fonts.
Use docker compose up --build for the bundled local Linux toolchain. Only host
127.0.0.1:5173 is published; /engine and /ingest proxy to internal loopback
services and preserve native /api paths and streaming. /wasm hosts Docxodus assets.

The launcher identifies an existing runtime owner. Preserve an existing stack
unless its owner has explicitly agreed to a restart. Never share one runtime
between simultaneous engines.
