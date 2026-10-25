import { expect, test } from "bun:test"
import { htmlBodyText } from "../lib/email"

// The tag strip an email's HTML-only body has always had. Stored citation
// offsets into existing emails depend on its output, so the linear version
// must give the same string for every input.
const regexStrip = (html: string) => html.replace(/<[^>]+>/g, " ")

test("the HTML tag strip gives the same text as the regex it replaced", () => {
  const bodies = [
    "",
    "Plain text, no tags.",
    "<p>Dear Recipient,</p>\r\n<p>The terms are <b>agreed</b>.</p>",
    '<html><head><style>p { margin: 0 }</style></head><body><p class="x">Fee &amp; costs</p><img src="cid:logo@example.test"></body></html>',
    "<!-- a comment --><br/><br />Line<>after<a <b>> c",
    "1 < 2 and 3 > 2, <<x>> <> <",
    "trailing <unclosed",
    ">leading close",
  ]
  // Generated inputs over the characters that matter: every arrangement of
  // tags, empty tags, nested and unclosed brackets the strip can meet.
  let seed = 7
  const random = () => (seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31
  const alphabet = ["<", ">", "a", " ", "\n", "<>", "</p>", "é"]
  for (let count = 0; count < 2000; count++)
    bodies.push(Array.from({ length: Math.floor(random() * 40) }, () => alphabet[Math.floor(random() * alphabet.length)]).join(""))
  for (const body of bodies) expect(htmlBodyText(body)).toBe(regexStrip(body))
})

test("an HTML body full of unclosed brackets is stripped in linear time", () => {
  // The regex took over a second at 100,000 brackets and minutes at a million.
  const started = performance.now()
  expect(htmlBodyText("<".repeat(300_000))).toBe("<".repeat(300_000))
  expect(htmlBodyText(`${"<a".repeat(150_000)}>`)).toBe(" ")
  expect(performance.now() - started).toBeLessThan(1_000)
})
