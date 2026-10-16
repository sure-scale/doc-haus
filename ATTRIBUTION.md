# Attribution

doc.haus stands on open work. This file credits the projects and datasets it
builds on or draws from. See [LICENSING.md](LICENSING.md) for how doc.haus itself
is licensed.

## OpenCode (MIT)

doc.haus is a fork of [OpenCode](https://github.com/anomalyco/opencode),
Copyright (c) 2025 opencode, licensed under the MIT License. The agent engine —
sessions, the multi-agent harness, permissions, the skill and tool system, and
model-provider plumbing — is OpenCode's work; doc.haus retargets it onto legal
documents. OpenCode's MIT license is preserved in [LICENSE.opencode](LICENSE.opencode).
The combined doc.haus release is licensed under GPLv3; see
[LICENSING.md](LICENSING.md). doc.haus is
not affiliated with or endorsed by the OpenCode team.

## Claude for Legal (Apache-2.0) — design inspiration

Some of doc.haus's legal skills are **inspired by the design and architecture** of
[anthropics/claude-for-legal](https://github.com/anthropics/claude-for-legal),
Copyright 2026 Anthropic PBC, licensed under the Apache License, Version 2.0.

This is a clean-room distillation: doc.haus drew on concepts and structure (triage
models, the "skills carry process, configuration carries positions" pattern,
treating gaps as first-class outputs) and wrote all of its own skill text
independently. No source text from claude-for-legal was copied into doc.haus. We
credit it here as a matter of good faith, not legal obligation, and doc.haus is
not affiliated with or endorsed by Anthropic. If any claude-for-legal text is ever
incorporated verbatim in future, that file will carry the Apache-2.0 license and a
statement of changes as Section 4 of that license requires.

See [docs/research/claude-for-legal-patterns.md](docs/research/claude-for-legal-patterns.md)
for the recorded provenance.

## CUAD — Contract Understanding Atticus Dataset (CC BY 4.0)

doc.haus's contract-risk checklist reproduces clause-category names from the
[Contract Understanding Atticus Dataset](https://www.atticusprojectai.org/cuad) by
The Atticus Project, licensed under
[Creative Commons Attribution 4.0 International](https://creativecommons.org/licenses/by/4.0/). The mapping
of those categories to doc.haus's checklist areas is doc.haus's own. Attribution
also appears inline at
`dochaus/skill/contract-risk-checklist/references/cuad-taxonomy.md`.

## Common Paper agreements (CC BY 4.0)

The [Common Paper Data Processing Agreement](https://github.com/CommonPaper/DPA)
by Common Paper supplies positions and mechanics used in
[the DPA playbook](dochaus/playbooks/playbook-dpa/SKILL.md). The
[Common Paper Cloud Service Agreement](https://commonpaper.com/standards/cloud-service-agreement/)
by Common Paper supplies market-norm baselines used in
[the SaaS/MSA playbook](dochaus/playbooks/playbook-saas-msa/SKILL.md).
Both sources are licensed under
[Creative Commons Attribution 4.0 International](https://creativecommons.org/licenses/by/4.0/).

doc.haus adapts these sources into executable review playbooks: positions are
organized as preferred terms, fallbacks, unacceptable terms and rationale;
clause wording and review guidance are adapted to each playbook's stated scope.
These playbooks are modified material, not unchanged copies of the agreements.
The source material retains its CC BY 4.0 licence. No Common Paper endorsement
is implied.

## Bonterms Cloud Terms (CC BY 4.0)

The [Bonterms Cloud Terms, Version 1.0](https://github.com/Bonterms/Cloud-Terms)
by Bonterms supply market-norm baselines and the section references in
[the SaaS/MSA playbook](dochaus/playbooks/playbook-saas-msa/SKILL.md). The source
is licensed under
[Creative Commons Attribution 4.0 International](https://creativecommons.org/licenses/by/4.0/).

doc.haus adapts these baselines into preferred terms, fallbacks, unacceptable
terms and rationale, with customer-side and vendor-side review guidance and
adapted clause wording. This is modified material, not an unchanged copy of the
Cloud Terms. The source material retains its CC BY 4.0 licence. No Bonterms
endorsement is implied.

## Bonterms NDA playbook (public domain)

[The NDA playbook](dochaus/playbooks/playbook-nda/SKILL.md) adapts language from
the [Bonterms Playbook for the Bonterms Mutual NDA](https://bonterms.com/resources/nda-playbook).
Bonterms identifies that source as public domain and permits reuse and
modification without implying its endorsement. doc.haus conforms defined terms
to its Disclosing Party/Receiving Party vocabulary and incorporates the adapted
language into its own review positions, fallbacks and rationale. No Bonterms
endorsement is implied.
