const XMLNS = "http://www.w3.org/2000/xmlns/"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const fail = (): never => { throw Error("Unsupported table import XML") }

function canonical(node: Node): unknown {
  if (node.nodeType === 1) {
    const e = node as Element
    if (e.namespaceURI === MC) return fail()
    return [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.namespaceURI !== XMLNS)
      .map(a => {
        // Ignorable contains namespace prefixes, not ordinary string values.
        // Its lexical equality alone does not establish the same semantics.
        if (a.namespaceURI === MC) {
          if (a.localName !== "Ignorable") return fail()
          const bindings = a.value.trim().split(/\s+/).map(prefix => {
            const uri = e.lookupNamespaceURI(prefix)
            if (!uri) return fail()
            return [prefix, uri]
          })
          return [a.namespaceURI, a.localName, a.value, bindings]
        }
        if (a.namespaceURI === "http://www.w3.org/2001/XMLSchema-instance" && a.localName === "type") return fail()
        return [a.namespaceURI, a.localName, a.value]
      }).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    Array.from(e.childNodes).map(canonical).filter(n => n !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) return node.nodeValue?.trim() || ["t", "delText", "binaryData"].includes((node.parentNode as Element)?.localName) ? ["text", node.nodeValue] : undefined
  return [node.nodeType, node.nodeName, node.nodeValue]
}
export const sameTableImportXml = (a: Node, b: Node) => JSON.stringify(canonical(a)) === JSON.stringify(canonical(b))
