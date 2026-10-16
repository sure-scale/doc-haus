---
name: citation-verification
description: The rules for grounding every claim about a matter document in a verified citation — anchoring quotes with the cite tool, handling failed verification, and sourcing case law only from the case-law tool. Use whenever quoting, citing, summarizing, or making factual claims about a matter's documents or about legal authority.
---

A lawyer must never be handed a quote whose text does not exist in the
document, or a case that does not exist in the reporter. Every factual claim
about a matter document is anchored before it is presented; every case-law
claim comes from a live lookup. Memory is not a source.

<quote-first>
Work in the order quote → assert → retract, never assert → find a quote to fit.

- **Extract the supporting span before the claim.** For any sourced statement,
  first locate the verbatim span in the matter document (search-document /
  read-document) that supports it. If no retrieved span supports a point, the
  point has "No relevant quotes found" — do not make it.
- **Assert only from extracted quotes.** Every claim must trace to a span you
  actually pulled. Do not back-fill a quote to justify a conclusion you already
  wrote; the quote comes first and the claim follows from it.
- **Retract the unsupported.** Before presenting an answer, drop any claim that no
  extracted span supports rather than leaving it standing as if grounded.
- **Abstention is allowed and correct.** "I can't find this in the matter's
  approved sources" is a valid answer. Prefer it over an answer assembled from
  general legal knowledge — the matter's documents (and, for authority, the
  case-law tool) are the only source of truth; your training never substitutes
  for them or fills a gap they leave open.
</quote-first>

<anti-sycophancy>
- **Do not adopt the user's premise.** A request can carry a false legal or
  jurisdiction premise ("since this is governed by New York law ...", "the cap
  is $5M, so ..."). Verify the premise against the documents before relying on
  it, and answer the corrected question.
- **Surface and correct false premises** rather than quietly answering on top of
  them. Reporting that the documents contradict the premise is more useful than
  agreement.
- **Refuse to manufacture authority.** Never invent or recall a clause, section
  number, case, reporter citation, or governing-law fact to support an
  unsupported position or to agree with the user. An unbacked position stays
  unbacked and is reported as such.
</anti-sycophancy>

<anchoring>
- **Anchor every quotation.** Before any passage from a matter document goes
  into an answer, call `cite` with it. All four required args matter:
  `docPath` (the absolute path exactly as search-document or read-document
  returned it — never a guessed or reconstructed path), `documentName` (the
  human-readable name), `quote`, and `reason` (why the passage supports the
  claim), plus an honest integer `confidence` 1-5.
- **Quotes are verbatim, 10-600 characters.** Copy the text exactly as the
  document returned it: no paraphrase, no added or removed ellipses, no
  normalized punctuation or "cleaned up" spacing. Under 10 characters the tool
  rejects the call — extend the quote with its surrounding sentence. Over 600,
  split it into consecutive quotes or quote the operative sentence and put the
  surroundings in the optional `context` arg.
- **Claims without a quote are claims without support.** A statement like "the
  agreement has no cap on liability" must rest on cited text (the clause that
  governs liability, quoted) or be presented as unverified. Do not present a
  conclusion about a document's contents with no anchored passage behind it.
</anchoring>

<verification-ladder>
Every citation any tool returns — search-document hits and cite anchors alike —
is re-checked against the live file before you see it, on a three-step ladder:

1. **exact** — the stored span still reproduces the excerpt verbatim: verified.
2. **re-anchor** — the excerpt moved (the document was edited after indexing)
   but still appears; the span is repaired and marked re-anchored: verified,
   but the search index is stale.
3. **reject** — the excerpt is nowhere in the live document: the citation is
   dropped.

When a result reports re-anchored or rejected citations, relay the
recommendation to re-upload (re-ingest) the named document — search stays
inaccurate until then. Never quote or rely on a rejected passage.
</verification-ladder>

<failed-verification>
When `cite` cannot verify a quote, the quote is wrong — the tool re-extracted
the live document and the text is not in it.

- **Re-read, then re-quote.** Go back to `read-document` or `search-document`,
  find the passage, and copy the verified text verbatim into a new `cite` call.
- **Never paraphrase-and-retry.** Reworking the failed quote from memory —
  shuffling words, trimming clauses, "fixing" it until something passes — is
  manufacturing a fabricated quotation. The only acceptable retry is text
  freshly copied from a tool result.
- **If the passage genuinely is not in the document**, say so. "I could not
  verify this in [document]" is a correct answer; an unanchored quotation
  never is.
</failed-verification>

<case-law>
- Case names, reporter citations, holdings, and quotations from opinions come
  ONLY from the `case-law` tool's results (or a connected CourtListener
  source), never from memory — a remembered case is presumptively fabricated.
  Cite only cases the tool returned, with the citation and URL it returned.
- The corpus is U.S. opinions only. On a non-U.S. matter, do not present its
  results as binding authority — flag them as U.S. precedent and route
  jurisdiction-specific authority questions to the supervising lawyer.
- When the tool returns nothing or is unreachable, report that. Never fill the
  gap with an invented or recalled citation.
- Auditing the citations in a finished brief or memo is its own procedure: the
  cite-check skill.
</case-law>
