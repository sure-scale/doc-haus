import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  Packer,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx"
import {
  chronologyExportDate,
  chronologyExportQualifications,
  chronologyExportEventNotes,
  chronologyExportCoverageNotes,
} from "../../../dochaus/lib/chronology-export-format"
import {
  CHRONOLOGY_EXPORT_ORIGIN,
  CHRONOLOGY_EXPORT_PROPERTY,
  type ChronologyExportSnapshot,
} from "../../../dochaus/lib/chronology-export"

// A4 portrait, one-inch margins; notes carry full accounts so table rows stay readable.
const PAGE = { width: 11906, height: 16838, margin: 1440 }
const WIDTHS = [1800, 4526, 2700]
const safe = (text: string) => text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "")
const unique = (values: string[]) => [...new Set(values.filter(Boolean))]
const paragraph = (text: string, options: { bold?: boolean; size?: number; color?: string } = {}) =>
  new Paragraph({
    spacing: { after: 100 },
    children: safe(text)
      .split(/\r\n|\n|\r/)
      .map((line, i) => new TextRun({ text: line, break: i ? 1 : undefined, ...options })),
  })
const heading = (
  text: string,
  level: typeof HeadingLevel.HEADING_1 | typeof HeadingLevel.HEADING_2 = HeadingLevel.HEADING_1,
) => new Paragraph({ heading: level, children: [new TextRun({ text: safe(text), color: "000000" })], keepNext: true })

/** Serialize only the authoritative, access-projected snapshot. No model call or source inference. */
export async function chronologyDocx(snapshot: ChronologyExportSnapshot): Promise<Buffer> {
  const cell = (index: number, lines: string[], header = false, eventCell = false) =>
    new TableCell({
      width: { size: WIDTHS[index], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 110, bottom: 110, left: 120, right: 120 },
      shading: header ? { type: ShadingType.CLEAR, fill: "EDEDED", color: "auto" } : undefined,
      children: lines.map((line, i) =>
        paragraph(line, {
          bold: header || (eventCell && i === 0),
          size: eventCell && i > 0 ? 18 : 20,
          color: eventCell && i > 0 ? "595959" : "000000",
        }),
      ),
    })
  const border = { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" }
  const table = new Table({
    width: { size: WIDTHS.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    columnWidths: WIDTHS,
    borders: {
      top: border,
      bottom: border,
      left: border,
      right: border,
      insideHorizontal: border,
      insideVertical: border,
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["Date", "Event", "Source"].map((text, i) => cell(i, [text], true)),
      }),
      ...snapshot.events.map(
        (event, i) =>
          new TableRow({
            cantSplit: true,
            children: [
              cell(0, [chronologyExportDate(event.date)]),
              cell(1, [`${i + 1}. ${event.title}`, chronologyExportQualifications(event).join("; ")], false, true),
              cell(
                2,
                unique(
                  event.accounts.flatMap((account) =>
                    account.evidence.map(
                      (evidence) => `${evidence.documentName}${evidence.section ? ` — ${evidence.section}` : ""}`,
                    ),
                  ),
                ).concat(event.accounts.some((account) => account.evidence.length) ? [] : ["Source not supplied"], [
                  `Evidence notes ${i + 1}`,
                ]),
              ),
            ],
          }),
      ),
    ],
  })
  const notes = snapshot.events.flatMap((event, i) => [
    heading(`${i + 1} ${event.title}`, HeadingLevel.HEADING_2),
    ...chronologyExportEventNotes(event).map(({ text, ...options }) => paragraph(text, options)),
  ])
  const document = new Document({
    creator: "doc.haus",
    title: safe(`Chronology ${snapshot.title}`),
    description: "Dated snapshot of saved chronology entries; derived work product, not independent source evidence.",
    customProperties: [{ name: CHRONOLOGY_EXPORT_PROPERTY, value: CHRONOLOGY_EXPORT_ORIGIN }],
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 22, color: "000000" },
          paragraph: { spacing: { after: 120, line: 276 } },
        },
      },
      paragraphStyles: [
        {
          id: "Title",
          name: "Title",
          basedOn: "Normal",
          run: { size: 38, bold: true, color: "000000" },
          paragraph: { spacing: { after: 220 }, keepNext: true },
        },
        {
          id: "Heading1",
          name: "heading 1",
          basedOn: "Normal",
          next: "Normal",
          run: { size: 28, bold: true, color: "000000" },
          paragraph: { spacing: { before: 260, after: 140 }, keepNext: true },
        },
        {
          id: "Heading2",
          name: "heading 2",
          basedOn: "Normal",
          next: "Normal",
          run: { size: 24, bold: true, color: "000000" },
          paragraph: { spacing: { before: 220, after: 100 }, keepNext: true },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE.width, height: PAGE.height },
            margin: { top: PAGE.margin, bottom: PAGE.margin, left: PAGE.margin, right: PAGE.margin },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({ text: "Chronology snapshot · ", size: 18 }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 18 }),
                ],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("Chronology")] }),
          paragraph(snapshot.title, { bold: true, size: 26 }),
          paragraph(
            `Snapshot ${new Date(snapshot.snapshotAt).toISOString()} · chronology revision ${snapshot.revision}`,
          ),
          paragraph(
            `${snapshot.events.length} of ${snapshot.totalEvents} saved entries. Scope: ${snapshot.scope === "filtered" ? "current filtered selection" : "all visible entries"}${snapshot.includeHidden ? ", with hidden entries explicitly included" : "; hidden entries excluded"}. Ordered earliest first; undated entries follow.`,
          ),
          paragraph(
            "This is derived chronology work product. Review status records review of an entry, not a finding that its account is true. Scheduled or required events do not establish that they occurred. Evidence notes and source coverage form part of this snapshot.",
          ),
          ...snapshot.warnings.map((warning) => paragraph(warning)),
          heading("Chronology entries"),
          ...(snapshot.events.length ? [table] : [paragraph("No saved entries match this selection.")]),
          heading("Evidence notes"),
          ...notes,
          heading("Source coverage"),
          paragraph(
            "Coverage describes chronology processing of matter sources at the snapshot time, including sources with no selected entry. Complete processing does not establish completeness or accuracy of the events.",
          ),
          ...snapshot.coverage.flatMap((coverage) =>
            chronologyExportCoverageNotes(coverage).map(({ text, ...options }) => paragraph(text, options)),
          ),
          ...(!snapshot.coverage.length ? [paragraph("No source coverage is recorded.")] : []),
        ],
      },
    ],
  })
  return Packer.toBuffer(document)
}
