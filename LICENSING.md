# Licensing

doc.haus v2 is **fully open source**. The doc.haus release as a whole, including
its legal layer, ingestion service and web frontend, is licensed under the
**GNU General Public License, version 3 only (GPL-3.0-only)**. The full license is
in [LICENSE](LICENSE), reproduced verbatim from
[GNU's official GPLv3 text](https://www.gnu.org/licenses/gpl-3.0.txt).

Copyright (c) 2026 doc.haus contributors (https://github.com/sure-scale/doc-haus).
The doc.haus legal components retain their copyright notice:
Copyright (c) 2026 SureScale.ai (doc.haus contributors).

## Scope and upstream provenance

doc.haus is a fork of [OpenCode](https://github.com/anomalyco/opencode).
OpenCode's original code remains available under its MIT License; this release
does not revoke or replace those permissions. Its copyright and full permission
notice are preserved in [LICENSE.opencode](LICENSE.opencode). The combined
doc.haus release is distributed under GPLv3, with third-party notices retained.

| Path | License | Notes |
| --- | --- | --- |
| doc.haus release as a whole | **GPL-3.0-only** | See [LICENSE](LICENSE). |
| `dochaus/` | **GPL-3.0-only**, except separately identified third-party material | Legal config layer (agents, skills, tools, prompts). See [dochaus/LICENSE](dochaus/LICENSE). |
| `services/engine/`, `services/ingest/` | **GPL-3.0-only** | doc.haus runtime and ingestion services. See [LICENSE](LICENSE) and [services/ingest/LICENSE](services/ingest/LICENSE). |
| `apps/web/` | **GPL-3.0-only** | Web frontend. See [apps/web/LICENSE](apps/web/LICENSE). |
| Original OpenCode code, including upstream packages | **MIT** | Upstream package metadata and notices are preserved; see [LICENSE.opencode](LICENSE.opencode). |

## What GPLv3 means for you

You may run, study, modify and share doc.haus, including commercially and in
client-facing services. If you distribute the software or modified versions,
follow GPLv3's requirements, including preserving notices, licensing covered
modifications under GPLv3 and providing the corresponding source code as required
by the license. The software is provided without warranty.

This is GPLv3, not the Affero GPL. Merely running a modified version as a network
service does not itself trigger a requirement to publish its source code.
Distributing browser code or other copies can still trigger GPLv3 obligations.

No additional grant is needed. [ADDITIONAL-GRANT.md](ADDITIONAL-GRANT.md) is retained
as a pointer for older links. This summary does not replace the full license.

## Third-party works, trademarks and attribution

Third-party dependencies and separately identified materials retain their own
licenses and notices. This release does not relicense those works. See
[ATTRIBUTION.md](ATTRIBUTION.md) and the notices accompanying those materials.

The "doc.haus" name and logo remain subject to the separate
[trademark policy](TRADEMARKS.md).
