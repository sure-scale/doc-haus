import { useEffect, useState } from "react"
import { listMatters, type Matter } from "../api/ingest/matters"

// Clients have no store of their own: a client is the Client value typed on
// one or more matters. Spellings that differ only in case or spacing are one
// client, shown under its most common spelling.

export type ClientGroup = { key: string; name: string; matters: Matter[] }

export function clientKey(name: string) {
  return name.replace(/\s+/g, " ").trim().toLowerCase()
}

export function groupClients(matters: Matter[]): ClientGroup[] {
  const groups = new Map<string, { matters: Matter[]; spellings: Map<string, number> }>()
  for (const matter of matters) {
    const spelling = matter.client?.replace(/\s+/g, " ").trim()
    if (!spelling) continue
    const key = spelling.toLowerCase()
    const group = groups.get(key) ?? { matters: [] as Matter[], spellings: new Map<string, number>() }
    group.matters.push(matter)
    group.spellings.set(spelling, (group.spellings.get(spelling) ?? 0) + 1)
    groups.set(key, group)
  }
  return [...groups].map(([key, group]) => {
    // Ties go to the spelling seen first, so the name does not flip between loads.
    let name = ""
    let best = 0
    for (const [spelling, count] of group.spellings) if (count > best) [name, best] = [spelling, count]
    return { key, name, matters: group.matters.sort((a, b) => b.created_at - a.created_at) }
  }).sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }))
}

// Names that start with what was typed come first, then names that contain it.
// An exact match is left out: picking it would change nothing.
export function suggestClients(names: string[], typed: string, limit = 8) {
  const term = clientKey(typed)
  if (!term) return names.slice(0, limit)
  const starts: string[] = []
  const contains: string[] = []
  for (const name of names) {
    const key = clientKey(name)
    if (key === term) continue
    if (key.startsWith(term)) starts.push(name)
    else if (key.includes(term)) contains.push(name)
  }
  return [...starts, ...contains].slice(0, limit)
}

// The client names already used on matters, loaded when `active` turns on (an
// edit opening) so suggestions include clients added since the page loaded.
export function useClientNames(active: boolean) {
  const [names, setNames] = useState<string[]>([])
  useEffect(() => {
    if (!active) return
    let live = true
    listMatters()
      .then((matters) => live && setNames(groupClients(matters).map((group) => group.name)))
      // Suggestions are a convenience; typing still works without them.
      .catch(() => undefined)
    return () => {
      live = false
    }
  }, [active])
  return names
}
