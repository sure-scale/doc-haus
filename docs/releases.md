# Product versions and releases

The default is one bundled product release at the end of a working day, when
there are changes ready to ship. Several feature PRs can merge into `dochaus`
during the day without changing the product version or publishing a release.
A release is a deliberate action, not a consequence of every feature merge.

## One version and one set of notes

`release.json` owns the product version (`name`), the increasing updater counter
(`version`) and the release notes (`notes`). The root `package.json` version must
match `name`. Independently versioned workspace packages and upstream components
keep their own versions. The app's updater already reads `release.json`.

Use stable semantic versions:

- **Patch**, such as 2.2.1: compatible bug fixes.
- **Minor**, such as 2.3.0: compatible new functionality, including any fixes
  shipped with it.
- **Major**, such as 3.0.0: breaking changes to supported interfaces, installation
  requirements or data compatibility. Document the migration.

Choose the highest level needed by the whole day's bundle. The date and number
of merged PRs do not determine the version. Prereleases are not supported by this
release channel. Do not bump the version in ordinary feature PRs.

## During the day

Merge reviewed feature PRs normally. Keep their user-facing change descriptions
clear enough to compile the end-of-day notes. Check all commits since the previous
release marker, including fixes and direct commits, so no shipped change is lost.
An optional draft can live outside the repository until release preparation.

Do not edit `release.json` to record work in progress, or to correct old notes:
the updater treats its latest changing commit as a release target. CI rejects a
changed marker without a higher semantic version and exactly one counter step.

## Prepare the daily release

Start a release branch from current reviewed `dochaus`, after the day's feature
PRs have landed. Write one reviewed Markdown notes file, for example
`/tmp/dochaus-release-notes.md`. Use concise user-facing bullets grouped by feature,
not one entry per PR. Include upgrade instructions or known limitations when
relevant. This file becomes the single source for the updater, GitHub Release
and the website entry.

```sh
bun run release:prepare minor --notes /tmp/dochaus-release-notes.md
bun run release:check --base origin/dochaus
git diff
```

`patch`, `major`, or an explicit higher stable version can replace `minor`.
Preparation updates `release.json` and the root `package.json`, and saves the
approved notes to `releases/<version>.md`. It does not commit, push, tag, publish
or deploy. Review the diff and open the release PR through the normal process.

Keep the release marker commit as the final content change in that PR. Rebase
onto the reviewed branch before preparing if other features have landed. If a
release PR needs additional changes, fold them into the release commit or prepare
it again from a fresh branch; do not create another counter increment for the
same release. CI requires the release marker and tested checkout to have the same
file tree, so the updater and GitHub tag identify the tested code.

## Publish and verify

Pushing or merging the release marker onto `dochaus` is the release action and
requires the owner's release instruction. Existing installations can see that
marker immediately: the current updater does not wait for GitHub Actions. Ensure
the release PR checks pass before merging. Branch protection should require
`release metadata` and the existing product CI checks, with branches required to
be up to date before merging. Otherwise a feature merged after the release PR
was tested can leave its marker behind the merged tree. The post-push check will
reject publication, but the updater marker will already be visible.

On that push, `dochaus-ci` validates metadata and runs its normal launcher,
Windows, typecheck, build and regression jobs. Only after all jobs succeed does
its publication job create `v<version>` at the exact updater release commit and
publish the GitHub Release from `releases/<version>.md`. Ordinary PRs, feature-only
pushes, other branches and manually dispatched checks do not publish releases.
Only the publication job has write permission. It refuses a mismatching existing
tag, or a release that is no longer the latest updater target. Re-running the same
successful publication is safe.

Verify the Actions result, GitHub Release, tag target and in-app offered version.
If publication fails, inspect the error and rerun the original failed workflow
once corrected. The updater marker may already be available; report that state
separately. If another feature push replaced a pending Actions run, rerun the
original release push run. Never move a published tag or reuse its version.
A newer release marker supersedes an older interrupted publication; reconcile
any historical missing release explicitly instead of pretending it is current.

Enable GitHub release immutability in repository settings before the first
publication where possible. This is a separate repository setting; adding the
workflow does not enable it. See [GitHub's immutable releases documentation](https://docs.github.com/en/code-security/concepts/supply-chain-security/immutable-releases).
The workflow publishes source releases; it does not build or claim to provide
standalone installer binaries.

## One website entry per day

In the separate website repository, create one entry for the release day using
the same approved `releases/<version>.md` notes. Display the date prominently and
the version alongside it, and link to the GitHub Release. Combine related work
into readable bullets rather than reproducing the PR list. Deploy the website
when authorized; product publication does not deploy it automatically.

A later urgent fix needs a new patch release and GitHub tag, but the website can
still keep one entry for that calendar day. Group the original release and hotfix
under that date with both version links and clearly identify the hotfix changes.
Do not rewrite the earlier version's notes to imply it included the hotfix.

## Initial adoption

The workflow setup aligns the root package version with the existing 2.2.0 marker
(counter 22), without changing that marker or publishing a release. Historical
markers without an archived Markdown file remain valid while unchanged. The next
release must have matching archived notes. Existing releases are not automatically
backfilled or renumbered. Backfill 2.2.0 only at its historical marker commit after
review and an explicit publication instruction, never at the current branch head.

For semantic-version rules, see [SemVer](https://semver.org/). For GitHub release
creation and tag verification, see the [GitHub CLI reference](https://cli.github.com/manual/gh_release_create).
