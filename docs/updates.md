# Updates and legacy migration

The in-app updater checks the public `origin` remote's `dochaus` branch and offers
a release only when release.json increases, including from worktrees and detached
checkouts. Feature-branch tracking settings never choose the update source.
Installing revalidates the exact offered release; a newer offer requires another
check rather than silently changing the install target. It requires a managed
launcher, no active work, a clean source tree and fast-forward ancestry. It installs frozen
dependencies and restores the previous source revision after installation
failure. Runtime settings, credentials, sessions and documents remain outside
source and survive successful updates and rollback. Keep backups independently.
The checked-out branch is not switched; local changes and separate code history
remain protected from updates.

## Existing installations with tracked settings

Older installations saved settings into dochaus/opencode.json. The updater
already installed there may reject a release changing that locally customized
file. It cannot be promised to bootstrap this storage change. Do not reset the
customized file or overwrite credentials/data to force an update.

Use an explicit administrator migration/install procedure:

1. Back up the legacy installation, its customized config and all data paths.
2. Prepare a separate checkout of the new candidate and install dependencies.
3. Run the utility with explicit absolute paths (dry-run is the default):

```sh
bun scripts/migrate-settings.ts --source /absolute/legacy/dochaus/opencode.json --runtime /absolute/installation/runtime
```

4. Review reported field names and unsupported customizations. No secret values
   are printed. Existing runtime overrides take precedence. Apply only after
   reviewing the dry run, using the same command with `--apply`.
5. The utility securely backs up the legacy configuration, validates supported
   mutable fields and publishes overrides atomically. Completion is recorded
   only after successful readback. Re-running is safe. Unsupported customizations
   remain in the legacy file/backup for deliberate review.
6. Launch the new checkout against the original workspace/runtime paths and
   verify credentials, settings, sessions and documents. Keep the old checkout
   and backup until verification succeeds. Return custom product-code changes
   separately; never import arbitrary plugin/MCP/instruction settings through
   the mutable settings API.

Migration does not move native credentials, databases or documents. This
preparation does not publish a bridge release or run migration on an existing
live installation.
