import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  BorderStyle,
  Footer,
  PageNumber,
  Table,
  TableRow,
  TableCell,
  WidthType,
} from "docx"
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync, copyFileSync } from "node:fs"
import path from "node:path"
import { createMatter, listMatters, WORKSPACE_ROOT } from "./matter"
import { ingestDocument } from "./ingest"
import { seedTemplates } from "./template"

// Two seeds live here, with different lifetimes:
//
//   seedPlaybooks() — the repo-shipped starter playbook library, copied into the
//   firm's WORKSPACE_ROOT/.playbooks on every ingest server boot (server.ts calls
//   it on startup). Not demo content: every install gets the starter set.
//
//   seedDemo() — the demo matter: a wholly fictional letter of engagement,
//   ingested through the real pipeline so a first-time user lands on a matter
//   that already answers cited questions and runs a legal review — no upload, no
//   data of their own required. Demo-gated: it only runs via this script
//   (`start.sh --demo` / `cd services/ingest && bun run seed`).
//
// seedDemo also writes the generated .docx to repo `demo/` so the same file can
// be dropped into a new matter through the web UI. Everything in it is invented;
// names and scenarios are fictional examples, not representations about real firms.

// dochaus/playbooks/ holds the repo-shipped starter playbooks — kept out of
// dochaus/skill/ so the engine and listPlaybooks() never auto-discover them from
// the repo; the firm's selectable library is WORKSPACE_ROOT/.playbooks, the same
// directory POST /playbooks imports into. Each starter is copied in at most once,
// tracked by name in the .seeded marker file: a starter the firm deleted stays
// deleted (delete-playbook is first-class), edits and same-name imports are never
// clobbered, and newly shipped starters still reach existing installs.
const PLAYBOOKS_SRC = path.join(import.meta.dir, "..", "..", "..", "dochaus", "playbooks")

export function seedPlaybooks() {
  const playbooksDir = path.join(WORKSPACE_ROOT, ".playbooks")
  const marker = path.join(playbooksDir, ".seeded")
  const seeded = new Set(existsSync(marker) ? readFileSync(marker, "utf8").split("\n").filter(Boolean) : [])
  const fresh = readdirSync(PLAYBOOKS_SRC)
    .filter((name) => name.startsWith("playbook-"))
    .filter((name) => existsSync(path.join(PLAYBOOKS_SRC, name, "SKILL.md")))
    .filter((name) => !seeded.has(name))
  if (!fresh.length) return
  fresh
    // A same-name import already present wins — record it as seeded without copying.
    .filter((name) => !existsSync(path.join(playbooksDir, name, "SKILL.md")))
    .forEach((name) => {
      mkdirSync(path.join(playbooksDir, name), { recursive: true })
      copyFileSync(path.join(PLAYBOOKS_SRC, name, "SKILL.md"), path.join(playbooksDir, name, "SKILL.md"))
      console.log(`Seeded starter playbook "${name}"`)
    })
  mkdirSync(playbooksDir, { recursive: true })
  writeFileSync(marker, [...seeded, ...fresh].join("\n") + "\n")
}

const DOC_NAME = "Letter of Engagement — Aldgate Mills.docx"
const MATTER_TITLE = "Aldgate Mills — Engagement (Demo)"

// The demo letter is an E&W law-firm engagement letter, so the matter binds the
// matching starter playbook — the full review pipeline (reviewer, playbook,
// challenger, summarizer) then demos end-to-end on first run.
const DEMO_PLAYBOOK = "playbook-engagement-letter"

// Each clause is a bold "N. Title" heading paragraph followed by its body. The
// ingest sectionizer keys sections off the leading clause number, so a citation
// like [Letter of Engagement — Aldgate Mills § 9] resolves to clause 9.
const CLAUSES: [string, string][] = [
  [
    "1. Scope of our work",
    "We will advise on and complete the acquisition of the long leasehold of Unit 5, Saffron Wharf, London E1, including reviewing the agreement for lease and the lease, reporting to you on title and on the principal commercial terms, raising and reviewing enquiries, and dealing with completion and post-completion registration at HM Land Registry. We will not advise on the commercial merits of the transaction, on tax beyond Stamp Duty Land Tax, or on the physical condition of the property, which are outside the scope of this engagement unless separately agreed in writing.",
  ],
  [
    "2. The people acting for you",
    "Your matter will be handled by Daniel Crane, Partner (charged at £480 per hour), with support from Priya Nair, Associate (charged at £290 per hour). Routine work may be delegated to a trainee or paralegal at £160 per hour where that is cost-effective. We will tell you promptly if the person responsible for your matter changes.",
  ],
  [
    "3. Our fees",
    "Our fees are calculated principally by reference to the time spent at the hourly rates in clause 2. Our current estimate for this matter is £14,500 plus VAT and disbursements. An estimate is not a fixed quotation; if it becomes likely that the estimate will be exceeded, we will tell you before further significant costs are incurred and agree a revised estimate with you.",
  ],
  [
    "4. Disbursements",
    "Disbursements are expenses we pay on your behalf, such as Land Registry fees, search fees, and Stamp Duty Land Tax. We will normally ask you to put us in funds for substantial disbursements before we incur them.",
  ],
  [
    "5. Billing and payment",
    "We will deliver interim bills monthly as the matter progresses, with a final bill on completion. Each bill is payable within 14 days of its date. We reserve the right to charge interest on bills not paid within that period at 4% per year above the base rate of the Bank of England, calculated from the date of the bill until payment.",
  ],
  [
    "6. Money on account",
    "We ask you to pay £5,000 on account of our fees and disbursements before we begin substantive work. We will hold that money in our client account and apply it against our bills, asking you to top it up as the matter proceeds.",
  ],
  [
    "7. Your responsibilities",
    "You agree to give us clear and timely instructions, to provide the documents and information we reasonably request, to put us in funds as agreed, and to tell us promptly of any change to the transaction or your instructions. Delay in any of these may affect our estimate and the timetable.",
  ],
  [
    "8. Confidentiality",
    "We will keep your affairs confidential, save where disclosure is required by law or by our regulator, or where you authorise it. Our duty of confidentiality continues after this engagement ends.",
  ],
  [
    "9. Limitation of liability",
    "Our total aggregate liability to you arising out of or in connection with this engagement, whether in contract, tort (including negligence), breach of statutory duty or otherwise, is limited to £3 million, which corresponds to our professional indemnity insurance cover. We are not liable for any indirect or consequential loss, or for loss of profit, revenue, or anticipated savings. Nothing in this clause limits any liability that cannot lawfully be limited, including liability for death or personal injury caused by negligence or for fraud.",
  ],
  [
    "10. Conflicts of interest",
    "We have checked for conflicts of interest and are not aware of any that prevent us acting for you. If a conflict arises during the engagement, we will tell you and explain the options, which may include our ceasing to act.",
  ],
  [
    "11. Data protection",
    "We process your personal data as a controller in order to act for you, in accordance with the UK GDPR and the Data Protection Act 2018. We retain your file for seven years after the matter closes, after which we may destroy it without further reference to you. Our privacy notice gives further detail.",
  ],
  [
    "12. Termination",
    "You may end this engagement at any time by written notice. We may cease to act only for good reason, such as a conflict of interest, non-payment of our bills, or your failure to give instructions, and on giving you reasonable written notice. On termination you remain liable for our fees and disbursements incurred up to that point.",
  ],
  [
    "13. Complaints",
    "We aim to provide a high standard of service. If you are unhappy with our service or a bill, please raise it first with Daniel Crane. If we cannot resolve it, you may be entitled to complain to the Legal Ombudsman, normally within six months of our final response. You may also have the right to challenge a bill under the Solicitors Act 1974.",
  ],
  [
    "14. Regulation",
    "Aldgate & Crane LLP is authorised and regulated by the Solicitors Regulation Authority. We are bound by the SRA Standards and Regulations, which are available from the SRA.",
  ],
  [
    "15. Governing law",
    "This engagement and our terms of business are governed by the law of England and Wales, and the courts of England and Wales have exclusive jurisdiction over any dispute arising out of them.",
  ],
  [
    "16. Acceptance",
    "If these terms are acceptable, please sign and date below and return one copy to us. Work you ask us to carry out, or the payment of money on account, will in any event be taken as your acceptance of these terms.",
  ],
]

// House palette for the (fictional) firm's letterhead.
const NAVY = "1C2B3A"
const GOLD = "C9A24B"
const INK = "222222"

// 22 half-points = 11pt body; clause/heading sizes follow. Rules are drawn as
// bottom paragraph borders so they print without a table.
const RULE = { bottom: { style: BorderStyle.SINGLE, size: 6, space: 4, color: GOLD } }

function body(text: string) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 160, line: 276 },
    children: [new TextRun(text)],
  })
}

function clause([title, text]: [string, string]) {
  const [num, ...rest] = title.split(". ")
  return [
    new Paragraph({
      spacing: { before: 220, after: 60 },
      children: [
        new TextRun({ text: `${num}.`, bold: true, color: GOLD }),
        new TextRun({ text: `  ${rest.join(". ")}`, bold: true, color: NAVY, allCaps: true, size: 20 }),
      ],
    }),
    body(text),
  ]
}

// Two-column block: recipient on the left, our ref / date on the right. A
// borderless table keeps the columns aligned the way a real letter sets them.
function metaCell(lines: string[], alignment: (typeof AlignmentType)[keyof typeof AlignmentType]) {
  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children: lines.map(
      (text, i) =>
        new Paragraph({
          alignment,
          spacing: { after: 20 },
          children: [new TextRun({ text, bold: i === 0, color: INK, size: 19 })],
        }),
    ),
  })
}

const NO_BORDERS = {
  top: { style: BorderStyle.NONE, size: 0, color: "auto" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
  left: { style: BorderStyle.NONE, size: 0, color: "auto" },
  right: { style: BorderStyle.NONE, size: 0, color: "auto" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
}

export async function seedDemo() {
  // The firm's emblem — a serif "A&C" monogram in a gold-ruled navy square. A
  // committed static asset, read at run time so the .docx carries a real
  // embedded image with no image-processing dependency.
  const logoPng = await Bun.file(path.join(import.meta.dir, "..", "assets", "logo.png")).bytes()

  const doc = new Document({
    styles: { default: { document: { run: { font: "Georgia", size: 22, color: INK } } } },
    sections: [
      {
        properties: { page: { margin: { top: 1100, bottom: 1100, left: 1300, right: 1300 } } },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                border: { top: { style: BorderStyle.SINGLE, size: 4, space: 6, color: GOLD } },
                alignment: AlignmentType.CENTER,
                spacing: { before: 60 },
                children: [
                  new TextRun({
                    text: "Aldgate & Crane LLP — fictional demonstration firm; not a registered practice.  ",
                    size: 14,
                    color: "888888",
                  }),
                  new TextRun({ text: "Authorised and regulated by the Solicitors Regulation Authority.", size: 14, color: "888888" }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 14, color: "888888" })],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 40 },
            children: [new ImageRun({ data: logoPng, type: "png", transformation: { width: 76, height: 76 } })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 20 },
            children: [new TextRun({ text: "ALDGATE & CRANE LLP", bold: true, color: NAVY, size: 30, allCaps: true })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            border: RULE,
            children: [new TextRun({ text: "S O L I C I T O R S", color: GOLD, size: 16 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 240 },
            children: [
              new TextRun({ text: "14 Saffron Court, London EC3N 4QX", size: 18, color: "555555" }),
              new TextRun({ text: "   ·   +44 (0)20 7946 0042   ·   law@aldgatecrane.example", size: 18, color: "555555" }),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: NO_BORDERS,
            rows: [
              new TableRow({
                children: [
                  metaCell(
                    ["Aldgate Mills Limited", "FAO: Ms R. Okafor, Director", "27 Wharf Road", "London E1 8GW"],
                    AlignmentType.LEFT,
                  ),
                  metaCell(["Our ref: A&C/2026-0042", "3 June 2026", "By email and post"], AlignmentType.RIGHT),
                ],
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 160 },
            children: [new TextRun({ text: "LETTER OF ENGAGEMENT", bold: true, color: NAVY, size: 26, allCaps: true })],
          }),
          body("Dear Ms Okafor,"),
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 160, line: 276 },
            children: [
              new TextRun({ text: "Re: Proposed acquisition of the long leasehold of Unit 5, Saffron Wharf, London E1. ", bold: true }),
              new TextRun(
                "Thank you for instructing Aldgate & Crane LLP. This letter sets out the basis on which we will act for you. Please read it, and let us know if anything is unclear, before signing and returning the acceptance at the end.",
              ),
            ],
          }),
          ...CLAUSES.flatMap(clause),
          new Paragraph({ spacing: { before: 240, after: 160 }, border: RULE, children: [] }),
          body("Yours sincerely,"),
          new Paragraph({
            spacing: { before: 200, after: 40 },
            children: [new TextRun({ text: "Daniel Crane", bold: true, color: NAVY })],
          }),
          body("Partner, for and on behalf of Aldgate & Crane LLP"),
          new Paragraph({
            spacing: { before: 240, after: 40 },
            children: [new TextRun({ text: "Signed (client):  ____________________________     Date:  ______________", color: INK })],
          }),
          body("Ms R. Okafor, for and on behalf of Aldgate Mills Limited"),
        ],
      },
    ],
  })

  const buffer = Buffer.from(withParagraphPointers(await Packer.toBuffer(doc)))

  // Keep a copy in repo demo/ so the same .docx can be uploaded through the UI.
  const demoDir = path.join(import.meta.dir, "..", "..", "..", "demo")
  mkdirSync(demoDir, { recursive: true })
  writeFileSync(path.join(demoDir, "Letter-of-Engagement-Aldgate-Mills.docx"), buffer)

  // Seed the firm's template library from repo demo/templates. Demo-only: a non-demo
  // boot ships an empty template library so first-run users start with nothing seeded.
  seedTemplates()

  // Idempotent: `start.sh --demo` runs this on every boot, so skip ingestion if the
  // demo matter is already present rather than piling up duplicates. The .docx above
  // is still rewritten so demo/ stays in sync with the seed script. The playbook is
  // bound only at creation — an existing matter's binding (including a deliberate
  // unbinding) is left alone.
  const existing = listMatters().find((m) => m.title === MATTER_TITLE)
  if (existing) {
    console.log(`Demo matter "${MATTER_TITLE}" already present (${existing.id}); skipping ingest.`)
    return
  }

  // The demo letter is an England & Wales engagement (SRA-regulated firm, UK GDPR,
  // E&W governing-law clause), so seed it with the EW jurisdiction pack — the first
  // run then shows jurisdiction-aware reasoning without any setup.
  const matter = createMatter(MATTER_TITLE, "A&C/2026-0042", ["EW"], DEMO_PLAYBOOK)
  const result = await ingestDocument(matter.dir, DOC_NAME, Buffer.from(buffer))
  console.log(`Seeded matter "${matter.title}" (${matter.id}), bound to ${DEMO_PLAYBOOK}`)
  console.log(`Ingested ${DOC_NAME}: ${result.sections} sections, ${result.chunks} chunks`)
  console.log(`Open the web app and select "${MATTER_TITLE}" to try cited Q&A and a legal review.`)
}

if (import.meta.main) {
  seedPlaybooks()
  await seedDemo()
}
