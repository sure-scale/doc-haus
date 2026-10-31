const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const elements = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const canon = (node: Node): unknown => {
  if (node.nodeType === 1) {
    const e = node as Element
    return [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.namespaceURI !== XMLNS).map(a => [a.namespaceURI, a.localName, a.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))), Array.from(e.childNodes).map(canon).filter(x => x !== undefined)]
  }
  return node.nodeType === 3 && !node.nodeValue?.trim() ? undefined : [node.nodeType, node.nodeValue]
}
const list = (part: Element) => {
  const data = elements(part), root = data.length === 1 && data[0]!.namespaceURI === PKG && data[0]!.localName === "xmlData" ? elements(data[0]!) : []
  const people = root.length === 1 && root[0]!.namespaceURI === W15 && root[0]!.localName === "people" ? root[0]! : undefined
  if (!people || Array.from(people.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === MC && a.localName === "Ignorable")) ||
      elements(people).some(e => e.namespaceURI !== W15 || e.localName !== "person")) return undefined
  return { shell: JSON.stringify(canon(people.cloneNode(false))), persons: elements(people) }
}
const person = (e: Element) => {
  const presence = elements(e)
  return Array.from(e.attributes).every(a => a.namespaceURI === XMLNS || a.namespaceURI === W15 && a.localName === "author") && presence.length === 1 &&
    presence[0]!.namespaceURI === W15 && presence[0]!.localName === "presenceInfo" && !elements(presence[0]!).length && Boolean(presence[0]!.getAttributeNS(W15, "providerId")) && Boolean(presence[0]!.getAttributeNS(W15, "userId")) &&
    Array.from(presence[0]!.attributes).every(a => a.namespaceURI === XMLNS || a.namespaceURI === W15 && ["providerId", "userId"].includes(a.localName))
}

/** Word lists each revision author in /word/people.xml. A first tracked change
 * by this user adds the part, or adds the user to a part that already lists
 * other reviewers. Every listed person stays as it was, in order; each added
 * person is one of this change's revision authors. */
export function nativeWordChatPeopleGrowth(before: Element | undefined, after: Element, authors: ReadonlySet<string>): boolean {
  const next = list(after), prior = before ? list(before) : { shell: next?.shell, persons: [] as Element[] }
  if (!next || !prior || prior.shell !== next.shell) return false
  const kept = prior.persons.map(e => JSON.stringify(canon(e))), known = new Set(prior.persons.map(e => e.getAttributeNS(W15, "author")))
  const added: Element[] = []
  let i = 0
  for (const e of next.persons) {
    if (i < kept.length && JSON.stringify(canon(e)) === kept[i]) { i++; continue }
    added.push(e)
  }
  const names = added.map(e => e.getAttributeNS(W15, "author") ?? "")
  return i === kept.length && (Boolean(before) || added.length > 0) && new Set(names).size === names.length &&
    added.every((e, k) => person(e) && authors.has(names[k]!) && !known.has(names[k]!))
}
