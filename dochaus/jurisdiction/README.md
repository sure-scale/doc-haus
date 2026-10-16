# Jurisdiction packs

A matter can carry several jurisdiction codes. Packs supply the legal reasoning
and citation context for substantive model stages. Add a directory beneath
`dochaus/jurisdiction/` with these files:

```text
<CODE>/
  profile.json
  prompt.md
```

Codes must start with a letter or digit and contain only letters, digits,
underscores or hyphens. The profile's `code` must match the directory name.

```json
{
  "code": "EW",
  "name": "England & Wales",
  "citationStyle": "OSCOLA",
  "preferredModel": null
}
```

`name` identifies the pack in the interface. `citationStyle` describes the
citation convention. `preferredModel` is currently descriptive metadata; it does
not select a provider, model or inference region. Model selection comes from
this installation's explicitly saved primary and fast settings.

Write `prompt.md` as concise guidance on governing law, the authority hierarchy,
citation conventions and any jurisdiction-specific reasoning limits. Ground the
guidance in general legal practice and preserve relevant provenance.

The ingest service exposes packs through `GET /jurisdictions` and stores the
user's selected codes in the matter's `matter.json` `jurisdictions` array.
[The pack loader](../lib/jurisdiction.ts) validates codes and loads their profiles
and prompt fragments. [Stage context](../plugin/legal/stage-context.ts) supplies
the selected packs to substantive stages; draft orchestration, routing and
narration receive no jurisdiction packs. [Legal policy](../plugin/legal/policy.ts)
adds each applicable pack as a `<jurisdiction>` system block. Missing prompt
files produce a warning and profile-only context; unknown or mismatched profiles
are not loaded.
