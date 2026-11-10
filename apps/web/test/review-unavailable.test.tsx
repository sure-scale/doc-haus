import { expect, test } from "bun:test"
import { renderToStaticMarkup } from "react-dom/server"
import { CompareSummary } from "../src/components/ReviewAttention"
import { readComparison } from "../src/api/review-compare"

test("a persisted unavailable comparison shows its limitation without agreement or difference labels", () => {
  const result = readComparison('{"status":"unavailable","reason":"A referenced exception is unavailable."}', { documents: [] })!
  const persisted = JSON.parse(JSON.stringify({ ...result, questionHash: "q", inputs: "evidence", comparedAt: 1 }))
  const html = renderToStaticMarkup(<CompareSummary comparison={persisted} onOpen={() => { throw Error("No document should be labelled.") }} />)
  expect(html).toContain("No AI comparison: A referenced exception is unavailable.")
  expect(html).not.toContain("agree")
  expect(html).not.toContain("differs")
})
