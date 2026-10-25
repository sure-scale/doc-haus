# Word redlines and document readings

A redline export is checked in both directions. Accepting all its tracked changes
must reproduce the intended document; rejecting them must reproduce the comparison
source. The check compares nonempty, normalized paragraph text in document order,
including notes and any headers/footers the comparison includes. It does not
certify visual layout or legal meaning. A mismatch refuses that export and names
the first differing paragraph; it does not change the source document.

If the comparison source already has pending tracked changes, the comparison uses
its accepted reading. Those earlier changes are not preserved as pending proposals
in the new redline. The source's revision count and authors appear in a native
Word comment and in the viewer, or in the drafting notes for a generated redline.
The original source remains unchanged.

Document tools provide three distinct Word readings for a planned document or an
exact registered uploaded Word filename:

- `current_body`: accepted text, keeping the current body as the reading of a
  generated document.
- `baseline_body`: reject-all text. For a generated document this reads its current
  redline when one exists.
- `marked_up_body`: a text projection with insertions, deletions and moves labelled
  inline with their recorded authors, including enclosing tracked rows and cells,
  also using the current generated redline.

The latter two are projections, not agreed text or current amendment targets.
They retain the source artifact name and byte identity, and use the same bounded
entry paging and lexical search as the other Word readings. Uploaded-source
indexing and citations remain separate from these projections.

Change note: previews previously checked accepted wording only when review comments
were attached, and current Word reads lacked baseline and markup projections.
Preview, drafting and version-compare exports now share the same text comparison;
read/search tools expose the separate readings. Focused regressions cover both
wrong export directions, nested text revisions/moves, source-revision disclosure,
and generated versus registered uploaded document access.
