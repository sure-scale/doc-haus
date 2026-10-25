# Reviewing drafting questions

Related questions now share topic navigation in the drafting question dialog.
Each remains a separate decision, with its own answer, advice and source passages.
Supporting passages can also explain their relevance to the question.

The dialog's Answered view keeps prior answers available, including adopted
recommendations. Change answer opens an explicit text correction; Save correction
uses the existing drafting answer queue. A failed submission keeps the words and
focus, and a queued correction is labelled until the coordinator applies it.
When no question remains open, Review answers opens the same dialog. Stopped
runs retain their record without offering corrections.

This extends the drafting dock, question card and shared dock shell; native
session question docks keep their existing behavior. The focused browser check
is `bun test/question-card.browser.ts` from `apps/web`; it covers independent
grouped decisions, failed and queued corrections, and review after the final
open question. `QUESTION_CARD_SCREENSHOTS` saves desktop and narrow captures.
