import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { decorateFilledValues } from "../src/components/document-markers"
import { FILLED_VALUE_FILL } from "../../../dochaus/lib/marker-patterns"

// Synthetic letter text; the class names stand in for the converter's.
test("highlights exactly the spans a batch preview shaded as filled values", () => {
  const { document } = new Window()
  document.body.innerHTML = `<style>.docx-a { background: #${FILLED_VALUE_FILL}; } .docx-b { background: #FFF2CC; }</style>
    <div id="page"><p>Dear <span class="docx-a">Example Contact</span>, re <span class="docx-b">[insert date]</span> and <span>plain text</span>.</p></div>`
  const page = document.getElementById("page") as unknown as HTMLElement
  expect(decorateFilledValues(page)).toBe(1)
  const marked = [...page.querySelectorAll(".doc-filled")]
  expect(marked.map((span) => span.textContent)).toEqual(["Example Contact"])
  expect(marked[0]!.getAttribute("title")).toBe("Filled in")
})
