import { expect, test } from "bun:test"
import { DEFAULT_DRAFTING, LEGEND_MAX_LENGTH as SERVER_LEGEND_MAX_LENGTH } from "../../../dochaus/lib/drafting-preferences"
import { NOTES_MARKING as SERVER_NOTES_MARKING, PAGE_NUMBER_AFFIXES, formatDraftDate as serverDate, pageFurniture } from "../../../dochaus/lib/page-furniture"
import { ADVICE_LEGENDS, AGREEMENT_LEGENDS, LEGEND_MAX_LENGTH, NOTES_MARKING, agreementFurniture, draftLineText, formatDraftDate, pageNumberText, suggestedLegends } from "../src/components/settings/page-layout"

test("legend suggestions follow the jurisdiction family and are empty elsewhere", () => {
  const us = { agreement: ["Draft — for discussion purposes only"], advice: ["Privileged & confidential — attorney-client communication"] }
  const commonwealth = { agreement: ["Subject to contract"], advice: ["Privileged & confidential"] }
  expect(suggestedLegends(["US"])).toEqual(us)
  expect(suggestedLegends(["US-NY", "US-DE"])).toEqual(us)
  for (const code of ["EW", "SCT", "NIR", "IE", "AU", "AU-NSW", "NZ", "SG", "HK"]) expect(suggestedLegends([code])).toEqual(commonwealth)
  for (const codes of [[], ["FR"], ["CA-ON"], ["EU"], ["ZA"]]) expect(suggestedLegends(codes)).toEqual({ agreement: [], advice: [] })
  expect(suggestedLegends(["EW", "US-CA", "IE"])).toEqual({ agreement: ["Subject to contract", "Draft — for discussion purposes only"], advice: ["Privileged & confidential", "Privileged & confidential — attorney-client communication"] })
})

test("presets offer no without-prejudice legend and every suggestion is a preset", () => {
  expect([...AGREEMENT_LEGENDS, ...ADVICE_LEGENDS].some(legend => /without prejudice/i.test(legend))).toBe(false)
  const all = suggestedLegends(["US", "EW"])
  expect(all.agreement.every(legend => (AGREEMENT_LEGENDS as readonly string[]).includes(legend))).toBe(true)
  expect(all.advice.every(legend => (ADVICE_LEGENDS as readonly string[]).includes(legend))).toBe(true)
  for (const legend of [...AGREEMENT_LEGENDS, ...ADVICE_LEGENDS]) expect(legend.length).toBeLessThanOrEqual(LEGEND_MAX_LENGTH)
})

test("the preview text matches the server's page furniture", () => {
  expect(LEGEND_MAX_LENGTH).toBe(SERVER_LEGEND_MAX_LENGTH)
  expect(NOTES_MARKING).toBe(SERVER_NOTES_MARKING)
  const at = new Date(2026, 9, 20, 12)
  for (const dateFormat of ["month-day-year", "day-month-year", "iso"] as const) expect(formatDraftDate(at, dateFormat)).toBe(serverDate(at.toISOString(), dateFormat)!)
  const formats = ["plain", "page", "dashed", "off"] as const
  let cases = 0
  for (const firm of ["Fixture Partners LLP", "  "])
    for (const draftLine of [true, false])
      for (const draftLineFirm of [true, false])
        for (const [index, pageNumbers] of formats.entries())
          for (const numberFirstPage of [true, false])
            for (const agreementLegend of ["", "Subject to contract"]) {
              const dateFormat = (["month-day-year", "day-month-year", "iso"] as const)[index % 3]!
              const pageNumberPosition = numberFirstPage ? "center" as const : "right" as const
              const prefs = { ...DEFAULT_DRAFTING, firm, dateFormat, draftLine, draftLineFirm, pageNumbers, pageNumberPosition, numberFirstPage, agreementLegend }
              const server = pageFurniture("agreement", prefs, at.toISOString())
              const web = agreementFurniture(prefs, at)
              expect(web.header).toEqual(server?.header ?? [])
              // Page 1's text, as the engine's footer prints it.
              expect(web.pageNumber?.text).toBe(server?.pageNumber && PAGE_NUMBER_AFFIXES[server.pageNumber.format].join("1"))
              expect(web.pageNumber?.align).toBe(server?.pageNumber?.align)
              expect(web.firstPage).toBe(server?.firstPage)
              cases++
            }
  expect(cases).toBe(128)
  expect(agreementFurniture({ ...DEFAULT_DRAFTING, firm: "Fixture Partners LLP", dateFormat: "day-month-year" }, at).header).toEqual(["Fixture Partners LLP draft: 20 October 2026"])
  expect(agreementFurniture(DEFAULT_DRAFTING, at)).toEqual({ header: ["Draft: October 20, 2026"], pageNumber: { text: "1", align: "center" } })
  expect(agreementFurniture({ ...DEFAULT_DRAFTING, pageNumbers: "dashed", numberFirstPage: false }, at)).toEqual({ header: ["Draft: October 20, 2026"], pageNumber: { text: "- 1 -", align: "center" }, firstPage: "unnumbered" })
  expect(draftLineText({ firm: "Fixture Partners LLP", draftLineFirm: false, dateFormat: "iso" }, at)).toBe("Draft: 2026-10-20")
  expect((["plain", "page", "dashed"] as const).map(format => pageNumberText(format, 2))).toEqual(["2", "Page 2", "- 2 -"])
})
