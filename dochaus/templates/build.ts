// Regenerates the checked-in drafting templates. Run from dochaus/:
//
//   bun templates/build.ts
//
// `_base.docx` is the seed every from-scratch draft starts from: an empty body
// carrying the house style set (Title, Heading 1-3, List Paragraph, Quote), so markdown
// inserted by Docxodus lands as real Word styles. It is assembled from raw OOXML
// parts zipped with the system `zip` (Docxodus can edit a .docx but not create
// one from nothing). `demo/templates/nda.docx` is then drafted from that seed
// through the same Docxodus insert pipeline the draft-document tool uses. It is
// demo content only: the ingest seed script copies it into the firm library
// (WORKSPACE_ROOT/.templates) when start.sh runs with --demo; a non-demo install
// starts with no templates.
//
// Template placeholder convention: every placeholder is a UNIQUE descriptive
// instruction, `[insert effective date]` — never a bare `[___]`. draft-document
// maps fills to placeholders by exact text, so two identical placeholders would
// both receive the same value (a date filled into "a period of [___] years").
import { mkdirSync, rmSync } from "node:fs"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { withParagraphPointers } from "../lib/drafting-amendment"
import { splitBlocks } from "../lib/markdown"

const here = path.dirname(fileURLToPath(import.meta.url))

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`

// Docxodus's HTML converter requires a settings part — without it conversion
// fails with "ArgumentNull_Generic ... part" — even though editing sessions
// open the document fine.
const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`

// One empty paragraph: Docxodus needs at least one body anchor to insert after.
const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p/>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="80"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="48"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="320" w:after="160"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="220" w:after="120"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:b/><w:i/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="720" w:right="720"/></w:pPr>
  </w:style>
</w:styles>`

const NDA_BODY = `# MUTUAL NON-DISCLOSURE AGREEMENT

This Mutual Non-Disclosure Agreement (this "Agreement") is entered into as of [insert effective date] (the "Effective Date") by and between [insert full legal name of first party], a [insert state] [insert entity type], and [insert full legal name of second party] (each a "Party" and together the "Parties"). A Party disclosing Confidential Information is the "Disclosing Party" and a Party receiving it is the "Receiving Party"; each Party may act in either capacity.

## 1. Purpose

The Parties wish to explore [insert description of the business opportunity] (the "Purpose"), and in connection with the Purpose each Party may disclose to the other certain confidential technical and business information that the Disclosing Party desires to be treated as confidential.

## 2. Confidential Information

"Confidential Information" means any information disclosed by either Party to the other Party, either directly or indirectly, in writing, orally, or by inspection of tangible objects, that is designated as confidential or that reasonably should be understood to be confidential given the nature of the information and the circumstances of disclosure. Confidential Information shall not include information that: (a) was publicly known prior to the time of disclosure; (b) becomes publicly known after disclosure through no action or inaction of the Receiving Party; (c) is already in the possession of the Receiving Party at the time of disclosure; (d) is obtained by the Receiving Party from a third party without a breach of any obligation of confidentiality; or (e) is independently developed by the Receiving Party without use of or reference to the Disclosing Party's Confidential Information.

## 3. Non-Use and Non-Disclosure

Each Party agrees not to use any Confidential Information of the other Party for any purpose except the Purpose. Each Party agrees not to disclose any Confidential Information of the other Party to third parties or to its employees, except to those employees, agents, and advisors who are required to have the information for the Purpose and who are bound by confidentiality obligations at least as protective as those in this Agreement.

## 4. Term

This Agreement shall remain in effect for a period of [insert term in years] years from the Effective Date. Each Party's obligations with respect to Confidential Information disclosed during the term shall survive for [insert survival period in years] years following expiration or termination, provided that obligations with respect to trade secrets shall survive for as long as the information remains a trade secret under applicable law.

## 5. Return of Materials

All documents and other tangible objects containing or representing Confidential Information shall be and remain the property of the Disclosing Party. Upon the Disclosing Party's written request, the Receiving Party shall promptly return or destroy all Confidential Information and certify such return or destruction in writing.

## 6. No License

Nothing in this Agreement grants any rights to either Party under any patent, copyright, trade secret, or other intellectual property right of the other Party, nor any rights to use Confidential Information except as expressly set out in this Agreement.

## 7. Remedies

Each Party acknowledges that any breach of this Agreement may cause irreparable harm for which monetary damages would be an inadequate remedy, and that the non-breaching Party shall be entitled to seek injunctive relief in addition to any other remedies available at law or in equity.

## 8. Governing Law

This Agreement shall be governed by and construed in accordance with the laws of [insert governing jurisdiction], without regard to its conflict of laws principles.

## 9. Non-Solicitation [optional: non-solicitation]

During the term of this Agreement and for [insert non-solicit period in months] months thereafter, neither Party shall directly or indirectly solicit for employment any employee of the other Party with whom it had contact in connection with the Purpose, provided that general solicitations not directed at such employees shall not breach this Section.

## 10. Entire Agreement

This Agreement constitutes the entire agreement between the Parties with respect to its subject matter and supersedes all prior agreements and understandings, whether written or oral. This Agreement may only be amended in a writing signed by both Parties.

IN WITNESS WHEREOF, the Parties have executed this Agreement as of the Effective Date.

[insert first party signature block]

[insert second party signature block]`

const stage = path.join(here, ".base-parts")
rmSync(stage, { recursive: true, force: true })
mkdirSync(path.join(stage, "_rels"), { recursive: true })
mkdirSync(path.join(stage, "word", "_rels"), { recursive: true })
await Bun.write(path.join(stage, "[Content_Types].xml"), CONTENT_TYPES)
await Bun.write(path.join(stage, "_rels", ".rels"), RELS)
await Bun.write(path.join(stage, "word", "document.xml"), DOCUMENT)
await Bun.write(path.join(stage, "word", "_rels", "document.xml.rels"), DOCUMENT_RELS)
await Bun.write(path.join(stage, "word", "styles.xml"), STYLES)
await Bun.write(path.join(stage, "word", "settings.xml"), SETTINGS)

const base = path.join(here, "_base.docx")
rmSync(base, { force: true })
const zip = Bun.spawnSync(["zip", "-q", "-X", "-r", base, "[Content_Types].xml", "_rels", "word"], { cwd: stage })
if (zip.exitCode !== 0) throw new Error(`zip failed: ${zip.stderr.toString()}`)
rmSync(stage, { recursive: true, force: true })
console.log(`built ${base}`)

const dx = await docxodus()
const session = dx.openDocxSession(await Bun.file(base).bytes(), {})
const seed = Object.keys(session.project().anchorIndex).find((id) => id.startsWith("p:"))!
const inserted = session.insertParagraph(seed, "after", splitBlocks(NDA_BODY))
if (!inserted.success) throw new Error(`nda insert failed: ${inserted.error?.message}`)
session.deleteBlock(seed)
const nda = path.join(here, "..", "..", "demo", "templates", "nda.docx")
mkdirSync(path.dirname(nda), { recursive: true })
await Bun.write(nda, withParagraphPointers(session.save()))
session.close()
console.log(`built ${nda}`)
