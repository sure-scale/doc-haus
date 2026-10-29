import React from "react"
import { createRoot } from "react-dom/client"
import { createBrowserRouter, RouterProvider } from "react-router-dom"
import App from "./App"
import "@fontsource/space-grotesk/500.css"
import "@fontsource/space-grotesk/600.css"
import "@fontsource/space-grotesk/700.css"
import "./styles.css"
import "./narration.css"
import "./narration-live.css"
import "./styles/primitives.css"
import "./styles/upload.css"
import "./styles/shell.css"
import "./styles/chat-log.css"
import "./styles/drafting.css"
import "./styles/drafting-items.css"
import "./styles/run-summary.css"
import "./styles/motion.css"
// Product design: foundations followed by the composed surface styles.
import "./styles/design-system.css"
import "./styles/design-shell.css"
import "./styles/design-documents.css"
import "./styles/word.css"
import "./styles/design-conversation.css"
import "./styles/design-library.css"
import "./styles/design-review.css"
import "./styles/design-settings.css"
import "./styles/design-matters.css"
import "./styles/launch-notice.css"
import "./styles/contrast.css"
import "./styles/modal-frost.css"

// A data router lets the review table stop same-document navigation while it
// has unsaved work. App retains its existing route tree and default basename.
const router = createBrowserRouter([{ path: "*", element: <App /> }])
createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    {/* URL-driven overlays (the review modal) must reopen in the same frame; the app has no Suspense to gain from transitions. */}
    <RouterProvider router={router} useTransitions={false} />
  </React.StrictMode>,
)
