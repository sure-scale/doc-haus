# Services

Run doc.haus with the root [launcher](../README.md). Its local services are:

| Service | Role | Entry point |
| --- | --- | --- |
| [engine](engine/) | Native OpenCode execution with the legal configuration boundary | bun run start |
| [ingest](ingest/) | Matter files, document parsing, retrieval and installation settings | bun run start |

The React product is in [apps/web](../apps/web/). Run package checks from the
owning directory. Ingest has its own frozen lockfile; a root workspace install
does not replace its package-local install and checks.

The retained [www](www/), [update](update/) and [files](files/) packages are
upstream OpenCode website and distribution tooling. Their READMEs describe
OpenCode infrastructure and publishing, not a configured doc.haus deployment.
The upstream deployment/publishing workflows are excluded from this release.
The root installer and [GitHub action](../github/) are also upstream components;
start.sh is the supported doc.haus launcher. doc.haus local updates use the
[installation updater](../docs/updates.md), not the upstream update Worker.
