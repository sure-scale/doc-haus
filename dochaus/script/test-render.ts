import { renderMarkdownToDocx } from "../lib/render.ts"

const source = Bun.argv[2]
if (!source) throw new Error("usage: bun script/test-render.ts <source.md>")

const content = await Bun.file(source).text()
console.log("File read. Size:", content.length)
const result = await renderMarkdownToDocx(content)
console.log("Render complete! Lint issues:", result.lint.length)
