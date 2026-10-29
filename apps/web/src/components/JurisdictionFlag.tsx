import { AE, AU, CA, CH, DE, EU, FR, GB, GB_SCT, HK, IE, IN, NL, NZ, SG, US, ZA } from "country-flag-icons/react/1x1"

// Flag artwork: country-flag-icons (MIT). Each jurisdiction pack shows its
// country's flag, and a state or province its country's with the region code
// beside it. England and Wales and Northern Ireland take the Union flag,
// Scotland the Saltire.
const FLAGS: Partial<Record<string, typeof GB>> = { AE, AU, CA, CH, DE, EU, FR, GB, HK, IE, IN, NL, NZ, SG, US, ZA, EW: GB, NIR: GB, SCT: GB_SCT }
const UK = new Set(["EW", "NIR", "SCT"])

/** The country a pack belongs to, as a browser region names it: EW is GB. */
export const countryOf = (code: string) => (UK.has(code.toUpperCase()) ? "GB" : code.toUpperCase().split("-")[0]!)

export default function JurisdictionFlag({ code, region = false }: { code: string; region?: boolean }) {
  const [base, sub] = code.toUpperCase().split("-")
  const Flag = FLAGS[base!]
  return (
    <span className="jx-flag" aria-hidden="true">
      {/* An unknown pack shows its letters, drawn by CSS so they stay out of the text. */}
      <span className={`jx-flag-disc${Flag ? "" : " is-blank"}`} data-code={Flag ? undefined : base!.slice(0, 2)}>{Flag && <Flag />}</span>
      {region && sub && <span className="jx-flag-region">{sub}</span>}
    </span>
  )
}
