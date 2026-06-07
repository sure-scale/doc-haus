// Generate a small legal-ish .docx for the spike: headings + numbered clauses,
// mirroring what ingest/mammoth will see (paragraphs separated by newlines).
import { Document, Packer, Paragraph, HeadingLevel, TextRun } from "docx"
import { writeFileSync } from "node:fs"

const clauses = [
  { h: "MASTER SERVICES AGREEMENT", body: "" },
  { h: "1. Definitions", body: 'In this Agreement, "Confidential Information" means any information disclosed by one party to the other.' },
  { h: "7. Termination", body: "Either party may terminate this Agreement for convenience upon thirty (30) days written notice to the other party." },
  { h: "7.2 Termination for Cause", body: "A party may terminate immediately if the other party commits a material breach that remains uncured for fifteen (15) days." },
  { h: "12. Governing Law", body: "This Agreement shall be governed by the laws of the State of Delaware, without regard to its conflict of laws principles." },
]

const children = []
for (const c of clauses) {
  children.push(new Paragraph({ text: c.h, heading: HeadingLevel.HEADING_1 }))
  if (c.body) children.push(new Paragraph({ children: [new TextRun(c.body)] }))
}

const doc = new Document({ sections: [{ children }] })
const buf = await Packer.toBuffer(doc)
writeFileSync("sample.docx", buf)
console.log("wrote sample.docx", buf.length, "bytes")
