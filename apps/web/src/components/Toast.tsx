import { createContext, useCallback, useContext, useState, type ReactNode } from "react"

// App-wide toasts: transient notifications stacked bottom-right. Type sets the
// left-border color — success/green, warning/amber, info/blue, error/red — over
// a white card with a soft shadow. Any component calls useToast() and fires one.
type ToastType = "success" | "warning" | "info" | "error"
type Toast = { id: number; type: ToastType; message: string }

const ToastContext = createContext<(type: ToastType, message: string) => void>(() => {})

export function useToast() {
  return useContext(ToastContext)
}

// Monotonic id source — unique within a session without Date.now().
let counter = 0

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([])
  const dismiss = useCallback((id: number) => setToasts((list) => list.filter((t) => t.id !== id)), [])
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
          <div key={t.id} className={`toast toast-${t.type}`} role="status">
            <span className="toast-msg">{t.message}</span>
            <button className="toast-close" onClick={() => dismiss(t.id)} title="Dismiss">
              <IconClose />
            </button>
          </div>
        ))}
      </div>
    </ToastContext.Provider>
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
