import { afterEach, beforeEach, expect, test } from "bun:test"
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

test("permission.asked projection preserves save scope, resources, and redaction metadata", () => {
  const scoped = projectEvent({
    type: "permission.asked", id: "event-permission", created: 1,
    data: {
      id: "permission-card", sessionID: "session-owner", action: "tracked-changes",
      resources: ["/matter/contract.docx"], save: ["/matter/*.docx", "/library/*.docx"], metadata: { document: "contract.docx" },
    },
  })[0]
  expect(scoped?.type).toBe("permission.asked")
  if (scoped?.type === "permission.asked") {
    expect(scoped.properties.always).toEqual(["/matter/*.docx", "/library/*.docx"])
    expect(scoped.properties.patterns).toEqual(["/matter/contract.docx"])
  }

  const onceOnly = projectEvent({
    type: "permission.asked", id: "event-once", created: 1,
    data: {
      id: "permission-once", sessionID: "session-owner", action: "redact",
      resources: ["/matter/personnel.docx"], save: [], metadata: { document: "personnel.docx", occurrences: 0, resave: true },
    },
  })[0]
  expect(onceOnly?.type).toBe("permission.asked")
  if (onceOnly?.type === "permission.asked") {
    expect(onceOnly.properties.always).toEqual([])
    expect(onceOnly.properties.metadata).toMatchObject({ document: "personnel.docx", occurrences: 0, resave: true })
  }
})
