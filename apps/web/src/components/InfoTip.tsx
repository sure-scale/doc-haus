import { Tooltip } from "./Tooltip"

// A small "i" badge that reveals a longer explanation on hover, so dense help
// text can sit beside a label instead of as a paragraph under every field.
export function InfoTip({ label }: { label: string }) {
  return (
    <Tooltip label={label}>
      <span className="info-tip" role="img" tabIndex={0} aria-label={label}>
        i
      </span>
    </Tooltip>
  )
}
