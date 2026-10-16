# One-command self-host: builds a fully-installed image, then `./start.sh` runs
# the three processes (engine + ingest + web) exactly as a local run does. This
# installs dependencies at build time and boots the managed stack at run time.
# `docker compose up` exposes the local web proxy on host loopback.
FROM denoland/deno:bin-2.8.2 AS deno
FROM oven/bun:1.4.2

# The shipped Python MCP server runs through Deno; match the locally verified version.
COPY --from=deno /deno /usr/local/bin/deno

# Baseline a real dev machine already has, that the bare bun image lacks:
#   - python3 + build-essential: the opencode monorepo's root install compiles
#     native node-gyp modules (e.g. tree-sitter-*).
#   - lsof + procps: start.sh detects occupied ports (lsof) and tears the
#     process tree down (pgrep) on exit.
#   - poppler-utils + tesseract-ocr + markitdown: scanned-document ingestion
#     and PDF fallback support. services/ingest uses LiteParse from its Bun
#     dependencies for primary PDF -> Markdown extraction, shells out to
#     markitdown if LiteParse fails, and to pdftoppm + tesseract (OCR for flat
#     scans) when they are on PATH, so the self-host image bakes the fallbacks in.
#   - libreoffice-writer + Carlito + fontconfig: headless DOCX rendering for
#     render-usability inspection, with a metric-compatible Calibri substitute
#     and deterministic font resolution. Poppler supplies the PDF inspection
#     commands (pdfinfo, pdftotext, and pdffonts).
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip build-essential lsof procps \
    poppler-utils tesseract-ocr libreoffice-writer \
    fonts-crosextra-carlito fontconfig \
  && rm -rf /var/lib/apt/lists/* \
  && pip3 install --no-cache-dir --break-system-packages "markitdown[pdf]"

WORKDIR /app
COPY . .

# Both dependency graphs use committed locks and patches.
RUN bun install --frozen-lockfile \
  && bun install --cwd services/ingest --frozen-lockfile \
  && bun run --cwd apps/web build

# Engine and ingest remain loopback-only; the web process proxies both.
EXPOSE 5173
ENV SKIP_INSTALL=1 NO_OPEN=1 DOCHAUS_DATA_ROOT=/data \
    OPENCODE_HOSTNAME=127.0.0.1 INGEST_HOST=127.0.0.1 WEB_HOST=0.0.0.0 \
    VITE_OPENCODE_URL=/engine VITE_INGEST_URL=/ingest \
    DOC_HAUS_ENGINE_PORT=4096 DOC_HAUS_INGEST_PORT=4500 DOC_HAUS_WEB_PORT=5173
CMD ["./start.sh"]
