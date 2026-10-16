import { afterEach, beforeEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import PermissionCard from "../src/components/chat/PermissionCard"
import { autoApproved, DEFAULT_PREFS, loadPrefs, savePrefs } from "../src/prefs"
import { projectEvent } from "../src/api/session-view"

const windowDescriptor = Object.getOwnPropertyDescriptor(globalThis, "window")
const documentDescriptor = Object.getOwnPropertyDescriptor(globalThis, "document")
let saved: string | null
let writable: boolean

beforeEach(() => {
  saved = null
  writable = true
  Object.defineProperty(globalThis, "window", {
    configurable: true,
    value: {
      localStorage: {
        getItem: () => saved,
        setItem: (_key: string, value: string) => {
          if (!writable) throw new Error("Storage unavailable")
          saved = value
        },
      },
    },
  })
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: { documentElement: { dataset: {} } },
  })
})

afterEach(() => {
  if (windowDescriptor) Object.defineProperty(globalThis, "window", windowDescriptor)
  if (!windowDescriptor) delete (globalThis as { window?: unknown }).window
  if (documentDescriptor) Object.defineProperty(globalThis, "document", documentDescriptor)
  if (!documentDescriptor) delete (globalThis as { document?: unknown }).document
})

const proposals = ["tracked-changes", "redline"]
const creations = ["draft-document", "create-template"]
const otherActions = ["word-integration", "redact", "update-template", "delete-template", "question", "unknown-action"]

test("approval choices default to ask and each switch covers only its own actions", () => {
  expect(loadPrefs()).toEqual(DEFAULT_PREFS)
  for (const action of [...proposals, ...creations, ...otherActions]) expect(autoApproved(action)).toBe(false)

  savePrefs({ autoApproveRedlines: true })
  for (const action of proposals) expect(autoApproved(action)).toBe(true)
  for (const action of [...creations, ...otherActions]) expect(autoApproved(action)).toBe(false)

  savePrefs({ autoApproveRedlines: false, autoApproveDrafting: true })
  for (const action of creations) expect(autoApproved(action)).toBe(true)
  for (const action of [...proposals, ...otherActions]) expect(autoApproved(action)).toBe(false)
})

test("saved boolean choices survive reads and disabling takes effect on the next request", () => {
  savePrefs({ autoApproveRedlines: true, autoApproveDrafting: true, textSize: "large", defaultJurisdictions: ["GB"] })
  expect(loadPrefs()).toMatchObject({ autoApproveRedlines: true, autoApproveDrafting: true })
  savePrefs({ autoApproveRedlines: false })
  expect(autoApproved("redline")).toBe(false)
  expect(autoApproved("draft-document")).toBe(true)
  expect(loadPrefs()).toMatchObject({ textSize: "large", defaultJurisdictions: ["GB"] })
})

test("only an explicit saved true enables automatic approval", () => {
  saved = JSON.stringify({ autoApproveRedlines: true, autoApproveDrafting: true })
  expect(autoApproved("redline")).toBe(true)
  expect(autoApproved("draft-document")).toBe(true)

  for (const value of ["false", "true", 1, 0, {}, [], null]) {
    saved = JSON.stringify({ autoApproveRedlines: value, autoApproveDrafting: value })
    expect(autoApproved("redline")).toBe(false)
    expect(autoApproved("draft-document")).toBe(false)
  }
  for (const value of ["invalid JSON", "null", "[]"]) {
    saved = value
    expect(autoApproved("redline")).toBe(false)
    expect(autoApproved("draft-document")).toBe(false)
  }
})

test("a required approval save reports storage failure and preserves the stored choice", () => {
  savePrefs({ autoApproveRedlines: true }, { requireStorage: true })
  expect(autoApproved("redline")).toBe(true)

  writable = false
  expect(() => savePrefs({ autoApproveRedlines: false }, { requireStorage: true })).toThrow("Could not save")
  expect(autoApproved("redline")).toBe(true)
})

test("existing callers can keep best-effort browser saves", () => {
  writable = false
  expect(() => savePrefs({ textSize: "large" })).not.toThrow()
  expect(loadPrefs().textSize).toBe("standard")
})

function permissionCard(save: string[]) {
  const event = projectEvent({
    type: "permission.asked", id: "event-permission", created: 1,
    data: {
      id: "permission-card", sessionID: "session-owner", action: "tracked-changes",
      resources: ["/matter/contract.docx"], save, metadata: { document: "contract.docx" },
    },
  })[0]
  if (event?.type !== "permission.asked") throw new Error("Native permission request was not projected")
  return renderToStaticMarkup(createElement(PermissionCard, { request: event.properties, onReply: () => {} }))
}

test("Always is available only when the native request supplies a save scope", () => {
  expect(permissionCard(["/matter/contract.docx"])).toContain("Always allow this action for this document")
  const onceOnly = permissionCard([])
  expect(onceOnly).toContain("Allow once")
  expect(onceOnly).toContain("Reject")
  expect(onceOnly).not.toContain("Always allow")
})

test("Always describes a broader native save scope rather than promising one document", () => {
  const html = permissionCard(["/matter/*.docx", "/library/*.docx"])
  expect(html).toContain("Always allow this action for these items")
  expect(html).toContain("/matter/*.docx")
  expect(html).toContain("/library/*.docx")
  expect(html).not.toContain("Always allow for this document")
})
