import { readBrowserValue, writeBrowserValue } from "../browser-storage"
import { useLayoutEffect, useRef, useState } from "react"
import logo from "../assets/dochaus-logo.svg"

// This is a per-tab acknowledgement, not an account-level consent record.
// Change the version when the notice changes so it is shown again.
const NOTICE_KEY = "dh.launch-notice.v1"

export function hasAcceptedLaunchNotice() {
  return readBrowserValue(NOTICE_KEY, true) === "accepted"
}

export default function LaunchNotice({ onAccept }: { onAccept: () => void }) {
  const [open, setOpen] = useState(true)
  const dialog = useRef<HTMLDialogElement>(null)
  const heading = useRef<HTMLHeadingElement>(null)

  useLayoutEffect(() => {
    if (open && !dialog.current?.open) {
      dialog.current?.showModal()
      heading.current?.focus()
    }
    if (!open && dialog.current?.open) dialog.current.close()
  }, [open])

  function accept() {
    writeBrowserValue(NOTICE_KEY, "accepted", true)
    onAccept()
  }

  return (
    <main className="launch-gate">
      <div className="launch-brand"><img src={logo} alt="" />Doc.Haus</div>
      {!open && (
        <section className="launch-paused">
          <h1>Doc.Haus is ready when you are.</h1>
          <p>Review and accept the AI use notice to continue.</p>
          <button className="primary" type="button" onClick={() => setOpen(true)}>Review notice</button>
        </section>
      )}
      <dialog
        ref={dialog}
        className="launch-notice"
        aria-labelledby="launch-notice-title"
        aria-describedby="launch-notice-intro"
        onCancel={(event) => { event.preventDefault(); setOpen(false) }}
      >
        <div className="launch-notice-body">
          <span className="launch-notice-eyebrow">Using Doc.Haus</span>
          <h1 ref={heading} id="launch-notice-title" tabIndex={-1}>Before you begin</h1>
          <p id="launch-notice-intro">Doc.Haus uses AI to assist with legal work. It does not provide legal advice or replace professional judgment.</p>
          <ul className="launch-notice-points">
            <li><strong>Review the output</strong><span>AI can make mistakes or miss important details. Review drafts and answers before using or sharing them.</span></li>
            <li><strong>Check the sources</strong><span>Verify citations, quotations and legal conclusions against current, relevant sources.</span></li>
            <li><strong>Protect client information</strong><span>Only add information you are authorised to use with your organisation’s configured AI services.</span></li>
          </ul>
          <p className="launch-notice-acknowledgement">By selecting “I accept”, you acknowledge these limitations and your responsibility to review the output.</p>
        </div>
        <div className="launch-notice-actions">
          <button type="button" onClick={() => setOpen(false)}>Not now</button>
          <button type="button" className="primary" onClick={accept}>I accept</button>
        </div>
      </dialog>
    </main>
  )
}
