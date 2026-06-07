# Docxodus spike scratch

Throwaway scripts backing the GO decision in `../docxodus-spike.md`. Run under
Bun against `docxodus@6.4.0`:

```
bun add docxodus mammoth docx   # docx only used to generate the sample
bun make-sample.mjs             # writes sample.docx
bun spike.mjs                   # Q1-Q4: init under Bun, html, offsets, tracked edit, redline
bun spike2.mjs                  # plain (untracked) edit path used by the word-integration tool
```
