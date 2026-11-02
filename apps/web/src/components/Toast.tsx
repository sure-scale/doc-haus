import { createContext, useCallback, useContext, useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react"
import { createPortal } from "react-dom"

// App-wide toasts: transient notifications stacked top-centre. Type sets the
// left-border color — success/green, warning/amber, info/blue, error/red — over
// a white card with a soft shadow. Any component calls useToast() and fires one.
type ToastType = "success" | "warning" | "info" | "error"
// settled: the entry animation has run, so moving between a dialog and the page never replays it.
type Toast = { id: number; type: ToastType; message: string; leaving?: boolean; settled?: boolean }

const ToastContext = createContext<(type: ToastType, message: string) => void>(() => {})

export function useToast() {
  return useContext(ToastContext)
}

// A native modal dialog (showModal) sits in the browser's top layer, above any z-index, and makes
// the rest of the page inert. Toasts fired while one is open render inside the topmost one so they
// are seen and announced; null means the page itself.
function topModal(): HTMLDialogElement | null {
  const open = [...document.querySelectorAll<HTMLDialogElement>("dialog[open]")]
  for (let index = open.length - 1; index >= 0; index--) {
    try {
      if (open[index].matches(":modal")) return open[index]
    } catch {
      return null
    }
  }
  return null
}

// Space to keep clear at the bottom of a dialog: its footer bar (marked data-toast-clear) holds the
// actions, such as Save, that a toast at the bottom would otherwise cover.
function clearance(dialog: HTMLDialogElement | null): number {
  const footer = dialog?.querySelector<HTMLElement>("[data-toast-clear]")
  if (!footer || footer.offsetHeight === 0) return 0
  return Math.max(0, window.innerHeight - footer.getBoundingClientRect().top)
}

// Monotonic id source — unique within a session without Date.now().
let counter = 0

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([])
  const [host, setHost] = useState<HTMLDialogElement | null>(null)
  const [lift, setLift] = useState(0)
  const follow = useCallback(() => {
    const modal = topModal()
    setHost(modal)
    setLift(clearance(modal))
  }, [])
  const timers = useRef(new Set<ReturnType<typeof setTimeout>>())
  useEffect(() => () => {
    for (const timer of timers.current) clearTimeout(timer)
    timers.current.clear()
  }, [])
  const schedule = useCallback((callback: () => void, delay: number) => {
    const timer = setTimeout(() => {
      timers.current.delete(timer)
      callback()
    }, delay)
    timers.current.add(timer)
  }, [])
  // Two-phase dismiss: mark the toast leaving so the exit transition plays,
  // then drop it once the animation has run.
  const dismiss = useCallback((id: number) => {
    setToasts((list) => list.map((t) => (t.id === id ? { ...t, leaving: true } : t)))
    schedule(() => setToasts((list) => list.filter((t) => t.id !== id)), 180)
  }, [schedule])
  // While toasts show, follow modal dialogs closing (close does not bubble, so listen in capture)
  // so a toast fired inside one moves back to the page instead of vanishing with it.
  const showing = toasts.length > 0
  useEffect(() => {
    if (!showing) return
    document.addEventListener("close", follow, true)
    return () => document.removeEventListener("close", follow, true)
  }, [showing, follow])
  const notify = useCallback(
    (type: ToastType, message: string) => {
      const id = ++counter
      follow()
      setToasts((list) => [...list, { id, type, message }])
      // Errors linger longer; they tend to carry something the user must read.
      schedule(() => dismiss(id), type === "error" ? 8000 : 5000)
    },
    [dismiss, schedule, follow],
  )

  return (
    <ToastContext.Provider value={notify}>
      {children}
      {createPortal(<div className="toast-stack" style={{ "--toast-lift": `${lift}px` } as CSSProperties}>
        {toasts.map((t) => (
          <div
            key={t.id}
            className={`toast toast-${t.type}${t.settled ? " toast-settled" : ""}${t.leaving ? " toast-leaving" : ""}`}
            role="status"
            onAnimationEnd={() => {
              if (!t.settled && !t.leaving) setToasts((list) => list.map((item) => (item.id === t.id ? { ...item, settled: true } : item)))
            }}
          >
            <span className="toast-icon">
              <ToastIcon type={t.type} />
            </span>
            <span className="toast-msg">{t.message}</span>
            <button type="button" className="toast-close" onClick={() => dismiss(t.id)} title="Dismiss" aria-label="Dismiss notification">
              <IconClose />
            </button>
          </div>
        ))}
      </div>, host?.isConnected ? host : document.body)}
    </ToastContext.Provider>
  )
}

function ToastIcon({ type }: { type: ToastType }) {
  const props = {
    width: 16,
    height: 16,
    viewBox: "0 0 24 24",
    fill: "none",
    stroke: "currentColor",
    strokeLinecap: "round",
    strokeLinejoin: "round",
  } as const
  if (type === "success")
    return (
      <svg {...props}>
        <circle cx="12" cy="12" r="10" />
        <path d="m8.5 12.5 2.5 2.5 5-6" />
      </svg>
    )
  if (type === "error")
    return (
      <svg {...props}>
        <circle cx="12" cy="12" r="10" />
        <path d="m15 9-6 6M9 9l6 6" />
      </svg>
    )
  if (type === "warning")
    return (
      <svg {...props}>
        <path d="m21.7 18-8-14a2 2 0 0 0-3.5 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.7-3Z" />
        <path d="M12 9v4" />
        <path d="M12 17h.01" />
      </svg>
    )
  return (
    <svg {...props}>
      <circle cx="12" cy="12" r="10" />
      <path d="M12 16v-4" />
      <path d="M12 8h.01" />
    </svg>
  )
}

function IconClose() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor">
      <line x1="18" y1="6" x2="6" y2="18" />
      <line x1="6" y1="6" x2="18" y2="18" />
    </svg>
  )
}
