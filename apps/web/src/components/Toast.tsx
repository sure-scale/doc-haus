import { createContext, useCallback, useContext, useState, type ReactNode } from "react"

// App-wide toasts: transient notifications stacked top-centre. Type sets the
// left-border color — success/green, warning/amber, info/blue, error/red — over
// a white card with a soft shadow. Any component calls useToast() and fires one.
type ToastType = "success" | "warning" | "info" | "error"
type Toast = { id: number; type: ToastType; message: string; leaving?: boolean }

const ToastContext = createContext<(type: ToastType, message: string) => void>(() => {})

export function useToast() {
  return useContext(ToastContext)
}

// Monotonic id source — unique within a session without Date.now().
let counter = 0

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([])
  // Two-phase dismiss: mark the toast leaving so the exit transition plays,
  // then drop it once the animation has run.
  const dismiss = useCallback((id: number) => {
    setToasts((list) => list.map((t) => (t.id === id ? { ...t, leaving: true } : t)))
    setTimeout(() => setToasts((list) => list.filter((t) => t.id !== id)), 180)
  }, [])
  const notify = useCallback(
    (type: ToastType, message: string) => {
      const id = ++counter
      setToasts((list) => [...list, { id, type, message }])
      // Errors linger longer; they tend to carry something the user must read.
      setTimeout(() => dismiss(id), type === "error" ? 8000 : 5000)
    },
    [dismiss],
  )

  return (
    <ToastContext.Provider value={notify}>
      {children}
      <div className="toast-stack">
        {toasts.map((t) => (
          <div key={t.id} className={`toast toast-${t.type}${t.leaving ? " toast-leaving" : ""}`} role="status">
            <span className="toast-icon">
              <ToastIcon type={t.type} />
            </span>
            <span className="toast-msg">{t.message}</span>
            <button type="button" className="toast-close" onClick={() => dismiss(t.id)} title="Dismiss" aria-label="Dismiss notification">
              <IconClose />
            </button>
          </div>
        ))}
      </div>
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
    strokeWidth: 2,
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
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <line x1="18" y1="6" x2="6" y2="18" />
      <line x1="6" y1="6" x2="18" y2="18" />
    </svg>
  )
}
