import ReactMarkdown from "react-markdown"
import { gfmTable } from "micromark-extension-gfm-table"
import { gfmTableFromMarkdown } from "mdast-util-gfm-table"
import type { Processor } from "unified"

// Extend the existing Markdown tree with tables only. Parsing the whole input
// keeps code, escapes, reference definitions and enclosing lists in one grammar.
function remarkTables(this: Processor) {
  const data = this.data()
  const syntax = data.micromarkExtensions ?? (data.micromarkExtensions = [])
  const trees = data.fromMarkdownExtensions ?? (data.fromMarkdownExtensions = [])
  syntax.push(gfmTable())
  trees.push(gfmTableFromMarkdown())
}

const plugins = [remarkTables]

export default function MarkdownContent({ children }: { children: string }) {
  return (
    <ReactMarkdown
      remarkPlugins={plugins}
      components={{
        table: ({ children }) => (
          <div className="markdown-table-scroll">
            <table className="markdown-table">{children}</table>
          </div>
        ),
      }}
    >
      {children}
    </ReactMarkdown>
  )
}
