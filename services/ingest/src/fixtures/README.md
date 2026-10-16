# Synthetic PDF regression fixtures

These four PDFs were freshly authored for doc.haus on 15 October 2026. All text,
page layouts, table cells and images originate in `generate-synthetic-pdfs.py`.
They contain invented workshop/equipment descriptions, not client documents,
actual parties, downloaded agreements, copied document wording or imported PDF
objects. The generator reads no external document. The scan embeds pixels from
our generated text-layer PDF; it has zero embedded text characters. Built-in PDF
Helvetica fonts are used without copied font files. These authored fixtures are
part of the doc.haus source release under the repository license.

| Fixture | Regression purpose |
| --- | --- |
| `synthetic-text-layer.pdf` | Born-digital text, title, numbered subheadings and complete body extraction. |
| `synthetic-continuation.pdf` | Two-page instruction continued across a page boundary, retained in its section and followed by a distinct next section. |
| `synthetic-table.pdf` | Bordered two-column form with distinct row labels, values and placeholder fields. |
| `synthetic-scanned.pdf` | Image-only copy of the authored workshop plan, exercising real Poppler/Tesseract OCR and the parser's scan fallback. |

Regenerate from this directory with Python 3, ReportLab 4.4.3, Pillow 11.3.0 and
Poppler `pdftoppm` on PATH. Use a disposable environment rather than installing
packages into the application checkout:

```sh
uv venv /tmp/dochaus-pdf-fixtures
uv pip install --python /tmp/dochaus-pdf-fixtures/bin/python reportlab==4.4.3 pillow==11.3.0
/tmp/dochaus-pdf-fixtures/bin/python generate-synthetic-pdfs.py
```

ReportLab uses invariant metadata. Repeated generation with the same toolchain
produces identical bytes. Poppler versions may rasterize fonts differently, so
record the rendering version when intentionally regenerating the scan. The
initial scan used Poppler 26.03.0. Metadata identifies only synthetic fixtures;
no user identity, original document metadata or local filesystem path is embedded.

Validation: render every page with `pdftoppm -png` and inspect the result; inspect
PDF metadata and verify the scan has no text layer; run `bun test src/pdf.test.ts`
from `services/ingest`. The tests retain strict structured-heading, table and OCR
assertions. They do not skip those checks when an extractor or OCR tool is absent.
Required runtime dependencies are the package's LiteParse/unpdf stack and the
local `pdftoppm` and `tesseract` executables. No model provider or paid request is
used. Regeneration tools are development-only and are not required to run tests
against the committed PDFs.
