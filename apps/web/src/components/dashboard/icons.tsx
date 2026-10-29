import type { ReactNode } from "react"

// Line icons in the app's style (24 grid, currentColor stroke).
function Icon({ children, size = 20 }: { children: ReactNode; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.75} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      {children}
    </svg>
  )
}

export const icons = {
  spend: <Icon><circle cx="12" cy="12" r="9" /><path d="M14.8 9.2c-.5-.9-1.6-1.4-2.8-1.4-1.6 0-2.8.8-2.8 2.1 0 3 5.6 1.4 5.6 4.3 0 1.3-1.2 2.1-2.8 2.1-1.3 0-2.4-.6-2.9-1.5M12 6.2v1.6M12 16.3v1.5" /></Icon>,
  matter: <Icon><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" /></Icon>,
  draft: <Icon><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" /><path d="M14 3v5h5M9 13h6M9 17h4" /></Icon>,
  documents: <Icon><path d="M8 7V5a2 2 0 0 1 2-2h6l4 4v10a2 2 0 0 1-2 2h-2" /><path d="M4 9a2 2 0 0 1 2-2h5l3 3v9a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2z" /></Icon>,
  week: <Icon><path d="M6 20V10M12 20V4M18 20v-7" /></Icon>,
  bars: <Icon><path d="M4 6h11M4 12h16M4 18h7" /></Icon>,
  attention: <Icon><path d="M6 8a6 6 0 1 1 12 0c0 7 3 8 3 8H3s3-1 3-8" /><path d="M10.3 20a1.9 1.9 0 0 0 3.4 0" /></Icon>,
  activity: <Icon><rect x="3" y="4" width="18" height="17" rx="3" /><path d="M3 9h18M8 2.5v3M16 2.5v3M7.5 13h.01M12 13h.01M16.5 13h.01M7.5 17h.01M12 17h.01" /></Icon>,
  recent: <Icon><circle cx="12" cy="12" r="9" /><path d="M12 7v5l3 2" /></Icon>,
  tokens: <Icon><path d="M12 3 3 8l9 5 9-5z" /><path d="m3 13 9 5 9-5" /></Icon>,
  trend: <Icon><path d="m3 17 6-6 4 4 8-8" /><path d="M15 7h6v6" /></Icon>,
  actions: <Icon><path d="M13 3 4 14h7l-1 7 9-11h-7z" /></Icon>,
  library: <Icon><path d="M4 19V5a2 2 0 0 1 2-2h12v14H6a2 2 0 0 0-2 2 2 2 0 0 0 2 2h12" /><path d="M9 7h5" /></Icon>,
  setup: <Icon><path d="M12 3l2.4 4.9 5.4.8-3.9 3.8.9 5.4L12 15.4 7.2 18l.9-5.4-3.9-3.8 5.4-.8z" /></Icon>,
  check: <Icon><path d="m5 12.5 4.2 4.2L19 7" /></Icon>,
  plus: <Icon><path d="M12 5v14M5 12h14" /></Icon>,
  arrowUp: <Icon size={16}><path d="M7 17 17 7M9 7h8v8" /></Icon>,
  arrowDown: <Icon size={16}><path d="M7 7l10 10M17 9v8H9" /></Icon>,
  arrowRight: <Icon size={16}><path d="M5 12h14M13 6l6 6-6 6" /></Icon>,
  chevronUp: <Icon size={16}><path d="m6 15 6-6 6 6" /></Icon>,
  chevronDown: <Icon size={16}><path d="m6 9 6 6 6-6" /></Icon>,
  hide: <Icon size={16}><path d="M3 3l18 18M10.6 10.6a2 2 0 0 0 2.8 2.8M9.4 5.2A9.7 9.7 0 0 1 12 5c5 0 9 5 9 7 0 .9-.8 2.4-2.2 3.8M6.1 6.1C4.2 7.5 3 9.6 3 12c0 2 4 7 9 7 1.7 0 3.3-.5 4.6-1.3" /></Icon>,
  reset: <Icon size={16}><path d="M4 4v6h6" /><path d="M4.6 15a8 8 0 1 0 1.8-8.3L4 10" /></Icon>,
  person: <Icon><circle cx="12" cy="8" r="4" /><path d="M4 21a8 8 0 0 1 16 0" /></Icon>,
  firm: <Icon><path d="M3 21h18M5 21V10l7-5 7 5v11M9 21v-6h6v6" /></Icon>,
  practice: <Icon><path d="M12 3v18M5 7h14M5 7l-3 7a3.5 3.5 0 0 0 6 0zM19 7l-3 7a3.5 3.5 0 0 0 6 0z" /></Icon>,
  globe: <Icon><circle cx="12" cy="12" r="9" /><path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18" /></Icon>,
  search: <Icon><circle cx="11" cy="11" r="7" /><path d="m20 20-3.5-3.5" /></Icon>,
  word: <Icon><rect x="3" y="4" width="18" height="16" rx="3" /><path d="m7 9 1.6 6L10.5 10l1.9 5L14 9M16.5 9h1" /></Icon>,
  style: <Icon><path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z" /><path d="m13.5 6.5 4 4" /></Icon>,
  model: <Icon><rect x="6" y="6" width="12" height="12" rx="3" /><path d="M9 2v4M15 2v4M9 18v4M15 18v4M2 9h4M2 15h4M18 9h4M18 15h4" /></Icon>,
  client: <Icon><circle cx="9" cy="8" r="3.5" /><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M21.5 20a6.5 6.5 0 0 0-4-6" /></Icon>,
  template: <Icon><rect x="4" y="3" width="16" height="18" rx="2" /><path d="M8 7h8M8 11h8M8 15h5" /></Icon>,
  playbook: <Icon><path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z" /><path d="m9 9 2 2 4-4" /></Icon>,
  sample: <Icon><path d="M9 3h6M10 3v6L4.5 18.5A1.7 1.7 0 0 0 6 21h12a1.7 1.7 0 0 0 1.5-2.5L14 9V3" /><path d="M7.5 14h9" /></Icon>,
  question: <Icon size={18}><circle cx="12" cy="12" r="9" /><path d="M9.5 9.3a2.6 2.6 0 0 1 5 .9c0 1.7-2.5 2.2-2.5 3.8M12 17h.01" /></Icon>,
  warning: <Icon size={18}><path d="M10.3 3.9 2.6 17.2A2 2 0 0 0 4.3 20h15.4a2 2 0 0 0 1.7-2.8L13.7 3.9a2 2 0 0 0-3.4 0z" /><path d="M12 9v4M12 17h.01" /></Icon>,
  pause: <Icon size={18}><circle cx="12" cy="12" r="9" /><path d="M10 9v6M14 9v6" /></Icon>,
  error: <Icon size={18}><circle cx="12" cy="12" r="9" /><path d="m9 9 6 6M15 9l-6 6" /></Icon>,
  plug: <Icon size={18}><path d="M9 2v6M15 2v6M6 8h12v4a6 6 0 0 1-12 0zM12 18v4" /></Icon>,
  flag: <Icon><path d="M5 21V4" /><path d="M5 4h12l-2.5 4.5L17 13H5" /></Icon>,
  idea: <Icon><path d="M9 18h6M10 21h4" /><path d="M12 3a6 6 0 0 0-3.6 10.8c.7.6 1.1 1.3 1.1 2.2h5c0-.9.4-1.6 1.1-2.2A6 6 0 0 0 12 3z" /></Icon>,
  review: <Icon><rect x="3" y="4" width="18" height="16" rx="2" /><path d="M3 10h18M9 10v10" /></Icon>,
  workflow: <Icon><circle cx="6" cy="6" r="2.5" /><circle cx="18" cy="18" r="2.5" /><path d="M8.5 6H15a3 3 0 0 1 0 6H9a3 3 0 0 0 0 6h6.5" /></Icon>,
  party: <Icon><path d="M4 20 8.5 7.5l8 8z" /><path d="M14 4.5c.5 1.5-.5 2.5-2 3M19.5 10c-1.5-.5-2.5.5-3 2M17 3v.01M21 7v.01M20 15v.01" /></Icon>,
} as const

export type IconName = keyof typeof icons

// Material 3 Expressive shapes as unit paths: a scalloped "cookie", a
// four-leaf "clover", a many-pointed "sunny" and a soft "pill" square.
function polar(lobes: number, depth: number, rotate = 0) {
  const points = Array.from({ length: 120 }, (_, index) => {
    const angle = (index / 120) * Math.PI * 2 + rotate
    const radius = 1 - depth + depth * Math.cos(lobes * (angle - rotate))
    return `${(Math.cos(angle) * radius).toFixed(3)} ${(Math.sin(angle) * radius).toFixed(3)}`
  })
  return `M${points.join("L")}Z`
}

export const SHAPES = {
  cookie: polar(9, 0.07),
  clover: polar(4, 0.16, Math.PI / 4),
  sunny: polar(12, 0.05),
  flower: polar(6, 0.11),
}

export type ShapeName = keyof typeof SHAPES

// An icon set in an expressive shape, filled with the badge's tone.
export function ShapeIcon({ icon, shape = "cookie", className = "" }: { icon: IconName; shape?: ShapeName; className?: string }) {
  return (
    <span className={`dash-shape ${className}`} aria-hidden="true">
      <svg viewBox="-1 -1 2 2" className="dash-shape-fill"><path d={SHAPES[shape]} /></svg>
      <span className="dash-shape-icon">{icons[icon]}</span>
    </span>
  )
}
