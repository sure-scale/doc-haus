import type { WordChatMaterial } from "../../../dochaus/lib/word-chat-contract"
import type { ChatAttachmentCatalog } from "./chat-attachments"
import type { DeskConnection } from "./desk-host"
import { CHAT_UPLOAD_MAX_BYTES, isChatUploadName } from "./desk-protocol"

const key = (material: { kind: string; id: string }) => `${material.kind}:${material.id}`
const reference = ({ kind, id, hash }: WordChatMaterial) => ({ kind, id, hash })
const icon = (path: string) => `<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${path}</svg>`
const fileIcon = icon('<path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z"/><path d="M14 3v5h5M8 12h8M8 16h5"/>')
const uploadIcon = icon('<path d="M12 16V3m-5 5 5-5 5 5M4 16v4a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-4"/>')

async function base64(file: File): Promise<string> {
  const bytes = new Uint8Array(await file.arrayBuffer())
  const chunks: string[] = []
  for (let start = 0; start < bytes.length; start += 32_768)
    chunks.push(String.fromCharCode(...bytes.subarray(start, start + 32_768)))
  return btoa(chunks.join(""))
}

export function mountAttachmentPicker(root: HTMLElement, connection: Pick<DeskConnection, "request">, options: { close(): void }) {
  root.classList.add("attachment-picker")
  root.setAttribute("aria-label", "Attach material to chat")
  root.innerHTML = `<header class="attachment-head"><div><h1>Attach material</h1><p data-picker="context"></p></div><button data-picker="close" type="button" aria-label="Close attachments" title="Close">×</button></header>
    <div class="attachment-tabs" role="tablist" aria-label="Attachment source">
      <button id="attachment-existing-tab" data-picker="existing-tab" type="button" role="tab" aria-controls="attachment-existing" aria-selected="true">${fileIcon}Existing</button>
      <button id="attachment-upload-tab" data-picker="upload-tab" type="button" role="tab" aria-controls="attachment-upload" aria-selected="false" tabindex="-1">${uploadIcon}Upload</button>
    </div>
    <section id="attachment-existing" data-picker="existing" role="tabpanel" aria-labelledby="attachment-existing-tab">
      <div class="attachment-search"><label class="attachment-sr" for="attachment-search">Search existing material</label><input id="attachment-search" data-picker="search" type="search" placeholder="Search documents and playbooks" autocomplete="off"><button data-picker="refresh" type="button" aria-label="Refresh material list" title="Refresh">↻</button></div>
      <ul class="attachment-list" data-picker="list" aria-label="Existing material"></ul><p data-picker="empty" class="attachment-empty" hidden></p>
    </section>
    <section id="attachment-upload" data-picker="upload" role="tabpanel" aria-labelledby="attachment-upload-tab" hidden>
      <div class="attachment-file"><span class="attachment-upload-icon">${uploadIcon}</span><label for="attachment-file">Choose a file</label><input id="attachment-file" data-picker="file" type="file" accept=".docx,.pdf,.txt,.md,.csv,.xlsx,.rtf" aria-describedby="attachment-formats"><p id="attachment-formats">Word, PDF, text, spreadsheet or RTF · up to 16 MB</p></div>
      <fieldset class="attachment-destination"><legend>Where to add it</legend><label><input data-picker="request-mode" type="radio" name="attachment-destination" value="request" checked><span>For this chat<small>Available in this conversation</small></span></label><label><input data-picker="matter-mode" type="radio" name="attachment-destination" value="matter"><span>Add to matter<small>Save with your matter documents</small></span></label></fieldset>
      <button data-picker="upload-button" class="attachment-primary" type="button" disabled>Upload &amp; attach</button>
    </section>
    <footer class="attachment-footer"><span data-picker="count"></span><button data-picker="done" type="button">Done</button></footer>
    <p data-picker="status" class="attachment-status" role="status" aria-live="polite"></p><p data-picker="error" class="attachment-error" role="alert" hidden></p>`
  const get = <T extends HTMLElement = HTMLElement>(name: string) => root.querySelector<T>(`[data-picker="${name}"]`)!
  const button = (name: string) => get<HTMLButtonElement>(name)
  const search = get<HTMLInputElement>("search"), file = get<HTMLInputElement>("file")
  const attached = new Map<string, WordChatMaterial>()
  let materials: WordChatMaterial[] = []
  let disposed = false, busy = false
  let view: "existing" | "upload" = "existing"
  function notice(text = "", error = false) {
    get("status").textContent = error ? "" : text
    get("error").textContent = error ? text : ""
    get("error").hidden = !error || !text
  }
  function controls() {
    root.setAttribute("aria-busy", String(busy))
    button("refresh").disabled = busy
    file.disabled = busy
    get<HTMLInputElement>("request-mode").disabled = busy
    get<HTMLInputElement>("matter-mode").disabled = busy
    button("upload-button").disabled = busy || !file.files?.length
    button("upload-button").textContent = busy && view === "upload" ? "Uploading…" : "Upload & attach"
    get("count").textContent = attached.size ? `${attached.size} attached` : "Choose material for your next message"
    for (const control of root.querySelectorAll<HTMLButtonElement>("[data-material]")) {
      const material = attached.get(control.dataset.material!)
      const selected = material?.hash === control.dataset.hash
      control.dataset.attached = String(selected)
      control.textContent = selected ? "✓ Attached" : "Attach"
      if (selected && material) control.setAttribute("aria-label", `Attached ${material.name}`)
      control.setAttribute("aria-disabled", String(busy || selected))
    }
  }
  function renderList() {
    const query = search.value.trim().toLocaleLowerCase()
    const visible = materials.filter(material => `${material.name} ${material.kind === "playbook" ? "playbook" : ""}`.toLocaleLowerCase().includes(query))
    const list = get("list")
    list.replaceChildren()
    for (const material of visible) {
      const item = document.createElement("li")
      const glyph = document.createElement("span"); glyph.className = "attachment-material-icon"; glyph.innerHTML = fileIcon
      const detail = document.createElement("div"); detail.className = "attachment-material-detail"
      const name = document.createElement("span"); name.textContent = material.name
      const type = document.createElement("small")
      type.textContent = material.kind === "playbook" ? "Playbook" : material.provenance === "request-upload" ? "For this chat" : "Matter document"
      detail.append(name, type)
      const choose = document.createElement("button"); choose.type = "button"; choose.dataset.material = key(material)
      choose.dataset.hash = material.hash
      const selected = attached.get(key(material))?.hash === material.hash
      choose.dataset.attached = String(selected)
      choose.textContent = selected ? "✓ Attached" : "Attach"
      choose.setAttribute("aria-label", `${selected ? "Attached" : "Attach"} ${material.name}`)
      choose.setAttribute("aria-disabled", String(busy || selected))
      choose.onclick = () => {
        if (busy || choose.dataset.attached === "true") return
        void work(async () => {
          const result = await connection.request({ action: "chat-attach", material: reference(material) }) as { material: WordChatMaterial }
          if (disposed) return
          attached.set(key(result.material), result.material)
          choose.dataset.attached = "true"; choose.textContent = "✓ Attached"
          choose.setAttribute("aria-label", `Attached ${result.material.name}`)
          notice(`${result.material.name} attached.`)
        })
      }
      item.append(glyph, detail, choose); list.append(item)
    }
    get("empty").hidden = visible.length > 0
    get("empty").textContent = query ? "No matching material." : "No material yet. Upload a file to get started."
    controls()
  }
  async function work(action: () => Promise<void>) {
    if (disposed || busy) return
    busy = true; notice(); controls()
    try { await action() }
    catch (error) { if (!disposed) notice(error instanceof Error ? error.message : "Could not attach this file. Try again.", true) }
    finally { busy = false; if (!disposed) controls() }
  }
  async function load() {
    await work(async () => {
      notice("Loading material…")
      const catalog = await connection.request({ action: "chat-materials" }) as ChatAttachmentCatalog
      if (disposed) return
      materials = catalog.materials
      attached.clear()
      for (const material of catalog.attached) attached.set(key(material), material)
      get("context").textContent = catalog.matterName
      renderList()
      notice()
    })
  }
  function show(next: typeof view, focus = false) {
    view = next
    for (const name of ["existing", "upload"] as const) {
      get(name).hidden = name !== next
      button(`${name}-tab`).setAttribute("aria-selected", String(name === next))
      button(`${name}-tab`).tabIndex = name === next ? 0 : -1
    }
    if (focus) button(`${next}-tab`).focus()
    controls()
  }
  for (const name of ["existing", "upload"] as const) {
    button(`${name}-tab`).onclick = () => show(name)
    button(`${name}-tab`).onkeydown = event => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return
      event.preventDefault()
      show(event.key === "Home" ? "existing" : event.key === "End" ? "upload" : view === "existing" ? "upload" : "existing", true)
    }
  }
  const close = () => { if (!disposed) options.close() }
  button("close").onclick = close
  button("done").onclick = close
  const keyboard = (event: KeyboardEvent) => { if (event.key === "Escape") { event.preventDefault(); close() } }
  root.addEventListener("keydown", keyboard)
  search.oninput = renderList
  button("refresh").onclick = () => { void load() }
  file.onchange = () => { notice(); controls() }
  button("upload-button").onclick = () => {
    const selected = file.files?.[0]
    if (!selected || busy) return
    if (!isChatUploadName(selected.name) || selected.size < 1 || selected.size > CHAT_UPLOAD_MAX_BYTES) {
      notice("Choose a supported file up to 16 MB.", true); return
    }
    const mode = get<HTMLInputElement>("matter-mode").checked ? "matter" : "request"
    void work(async () => {
      const data = await base64(selected)
      if (disposed) return
      const result = await connection.request({ action: "chat-upload", name: selected.name, data, mode }) as { material: WordChatMaterial }
      if (disposed) return
      attached.set(key(result.material), result.material)
      materials = [...materials.filter(material => key(material) !== key(result.material)), result.material]
      file.value = ""
      renderList()
      notice(mode === "matter" ? "Added to the matter and attached." : "Attached to this chat.")
    })
  }
  const ready = load().then(() => { if (!disposed) search.focus() })
  return { ready, dispose() {
    if (disposed) return
    disposed = true
    root.removeEventListener("keydown", keyboard)
    root.replaceChildren()
    root.classList.remove("attachment-picker")
    root.removeAttribute("aria-busy")
  } }
}
