/** Host alert/confirm dialogs can be suppressed in Office WebViews. Keep the
 * deliberate replacement choice inside the pane, bound to its displayed version. */
export function createReplaceConfirmation(document: Document) {
  let active: { close(): void } | undefined
  return {
    confirm(version: { number: number }): Promise<boolean> {
      if (active || !Number.isSafeInteger(version.number) || version.number < 1) return Promise.resolve(false)
      const previous = document.activeElement as HTMLElement | null
      const layer = document.createElement("div"); layer.className = "replace-confirmation-layer"
      layer.innerHTML = `<section class="replace-confirmation" role="alertdialog" aria-modal="true" aria-labelledby="replace-confirmation-title" aria-describedby="replace-confirmation-description" tabindex="-1">
        <h2 id="replace-confirmation-title"></h2>
        <p id="replace-confirmation-description">This replaces this Word copy, including unsaved changes. To keep a copy, cancel and use Save As first.</p>
        <p class="replace-confirmation-progress" role="status" aria-live="polite" hidden>Replacing contents…</p>
        <div class="replace-confirmation-actions"><button type="button" data-replace="cancel">Cancel</button><button type="button" data-replace="accept" class="primary">Replace contents</button></div>
      </section>`
      const dialog = layer.querySelector<HTMLElement>("section")!
      layer.querySelector("h2")!.textContent = `Load version ${version.number}?`
      const cancel = layer.querySelector<HTMLButtonElement>('[data-replace="cancel"]')!
      const accept = layer.querySelector<HTMLButtonElement>('[data-replace="accept"]')!
      let resolve!: (approved: boolean) => void
      const answer = new Promise<boolean>(done => { resolve = done })
      let decided = false, busy = false, closed = false
      const background = Array.from(document.body.children).map(element => ({ element, inert: element.getAttribute("inert"), hidden: element.getAttribute("aria-hidden") }))
      function close() {
        if (closed) return
        closed = true
        if (!decided) { decided = true; resolve(false) }
        document.removeEventListener("keydown", keydown, true)
        document.removeEventListener("focusin", focusin, true)
        layer.remove()
        for (const { element, inert, hidden } of background) {
          if (inert === null) element.removeAttribute("inert"); else element.setAttribute("inert", inert)
          if (hidden === null) element.removeAttribute("aria-hidden"); else element.setAttribute("aria-hidden", hidden)
        }
        active = undefined
        // The caller's Save queue/button finishes in subsequent microtasks.
        document.defaultView?.setTimeout(() => {
          if (previous?.isConnected && !previous.closest("[hidden], [inert]") && !previous.matches(":disabled")) previous.focus()
          else document.querySelector<HTMLElement>('.pane-tabs [role="tab"][aria-selected="true"]')?.focus()
        }, 0)
      }
      function focusin(event: FocusEvent) {
        if (!layer.contains(event.target as Node)) (busy ? dialog : cancel).focus()
      }
      function keydown(event: KeyboardEvent) {
        if (event.key === "Escape") {
          event.preventDefault(); event.stopPropagation()
          if (!busy) close()
        }
        if (event.key === "Tab") {
          event.preventDefault(); event.stopPropagation()
          if (busy) dialog.focus()
          else if (document.activeElement === cancel) accept.focus()
          else if (document.activeElement === accept) cancel.focus()
          else (event.shiftKey ? accept : cancel).focus()
        }
      }
      cancel.onclick = () => { if (!busy) close() }
      accept.onclick = () => {
        if (decided || busy) return
        decided = true; busy = true
        cancel.disabled = true; accept.disabled = true; accept.textContent = "Replacing…"
        dialog.setAttribute("aria-busy", "true")
        layer.querySelector<HTMLElement>(".replace-confirmation-progress")!.hidden = false
        dialog.focus(); resolve(true)
      }
      for (const { element } of background) { element.setAttribute("inert", ""); element.setAttribute("aria-hidden", "true") }
      document.body.append(layer)
      document.addEventListener("keydown", keydown, true)
      document.addEventListener("focusin", focusin, true)
      active = { close }
      cancel.focus()
      return answer
    },
    complete() { active?.close() },
    dispose() { active?.close() },
  }
}
