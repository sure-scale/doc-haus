const XMLNS = "http://www.w3.org/2000/xmlns/"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const names: Record<string, readonly string[]> = {
  [W]: ["document","body","tbl","tblPr","tblGrid","gridCol","tr","trPr","tc","tcPr","p","pPr","r","rPr","t","delText","sectPr","tcW","tblW","tblLayout","gridSpan","vMerge","hMerge","ins","del","styles","style","docDefaults","rPrDefault","sz","name","numPr","numId","ilvl","rsidR","rsidRDefault","rsidP","rsidDel","rsidSect","w","h","type","val","id","author","date","styleId","default"],
  [MC]: ["Ignorable","AlternateContent","Choice","Fallback"],
  "http://schemas.microsoft.com/office/word/2010/wordml": ["paraId","textId"],
  "http://schemas.microsoft.com/office/word/2023/wordml/word16du": ["dateUtc"],
  "http://schemas.microsoft.com/office/2006/xmlPackage": ["xmlData","binaryData"],
  "http://schemas.openxmlformats.org/package/2006/relationships": ["Relationships","Relationship"],
  "http://www.w3.org/XML/1998/namespace": ["space"],
}
const prefixes: Record<string,string> = {[W]:"w",[MC]:"mc","http://schemas.microsoft.com/office/word/2010/wordml":"w14","http://schemas.microsoft.com/office/word/2023/wordml/word16du":"w16du","http://schemas.microsoft.com/office/2006/xmlPackage":"pkg","http://schemas.openxmlformats.org/package/2006/relationships":"rel","http://www.w3.org/XML/1998/namespace":"xml"}
const safeName = (node: {namespaceURI: string | null;localName: string | null}, attribute = false) => {
  const uri=node.namespaceURI??"",name=node.localName??""
  if (Object.prototype.hasOwnProperty.call(names,uri) && names[uri]!.includes(name) && Object.prototype.hasOwnProperty.call(prefixes,uri)) return `${prefixes[uri]}:${name}`
  if (attribute && !uri && ["Id","Type","Target","TargetMode"].includes(name)) return name
  return attribute ? "attribute" : "element"
}
export type TableImportXmlDifference = {category:"element_name"|"attribute_presence"|"attribute_value"|"namespace_binding"|"text"|"node_count"|"node_kind"|"node_value";path:string}
type UnsupportedReason = "compatibility_element"|"compatibility_attribute"|"unbound_prefix"|"qualified_type"
export class TableImportXmlUnsupportedError extends Error {
  constructor(readonly path:string,readonly reason:UnsupportedReason) {super("Unsupported table import XML")}
}
const childPath = (path:string,node:Node,index:number) => `${path}/${node.nodeType===1?safeName(node as Element):"node"}[${index+1}]`
const bounded = (path:string) => path.split("/").slice(0,25).join("/")
const fail = (path:string,reason:UnsupportedReason): never => { throw new TableImportXmlUnsupportedError(bounded(path),reason) }
// A native body export can renew addition-session identifiers independently
// of the file export. This policy is used only for that live source preimage,
// never for oracle/capsule or post-import equality. All revision carriers remain.
const liveAdditionSessionAttribute = (element:Element,attr:Attr) => element.namespaceURI===W && attr.namespaceURI===W
  && ((["tr","p","r"].includes(element.localName) && attr.localName==="rsidR") || (element.localName==="p" && attr.localName==="rsidRDefault"))
  && /^[0-9a-fA-F]{8}$/.test(attr.value)

function canonical(node: Node,path = node.nodeType===1?`/${safeName(node as Element)}[1]`:"/node[1]",liveAdditionSession=false): unknown {
  if (node.nodeType === 1) {
    const e = node as Element
    if (e.namespaceURI === MC) return fail(path,"compatibility_element")
    return [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.namespaceURI !== XMLNS && !(liveAdditionSession && liveAdditionSessionAttribute(e,a)))
      .map(a => {
        // Ignorable contains namespace prefixes, not ordinary string values.
        // Its lexical equality alone does not establish the same semantics.
        if (a.namespaceURI === MC) {
          if (a.localName !== "Ignorable") return fail(`${path}/@${safeName(a,true)}`,"compatibility_attribute")
          const bindings = a.value.trim().split(/\s+/).map(prefix => {
            const uri = e.lookupNamespaceURI(prefix)
            if (!uri) return fail(`${path}/@mc:Ignorable`,"unbound_prefix")
            return [prefix, uri]
          })
          return [a.namespaceURI, a.localName, a.value, bindings]
        }
        if (a.namespaceURI === "http://www.w3.org/2001/XMLSchema-instance" && a.localName === "type") return fail(`${path}/@attribute`,"qualified_type")
        return [a.namespaceURI, a.localName, a.value]
      }).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    Array.from(e.childNodes).map((child,index)=>canonical(child,childPath(path,child,index),liveAdditionSession)).filter(n => n !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) return node.nodeValue?.trim() || ["t", "delText", "binaryData"].includes((node.parentNode as Element)?.localName) ? ["text", node.nodeValue] : undefined
  return [node.nodeType, node.nodeName, node.nodeValue]
}
export const sameTableImportXml = (a: Node, b: Node) => JSON.stringify(canonical(a)) === JSON.stringify(canonical(b))
export const sameLiveTableImportPreimageXml = (a:Node,b:Node) => JSON.stringify(canonical(a,undefined,true))===JSON.stringify(canonical(b,undefined,true))

/** Diagnostic only. Equality remains owned by the unchanged canonical predicate.
 * Names are expanded through a fixed vocabulary; no text, values, prefixes or
 * unknown element/attribute names are included in the returned path. */
function firstDifference(a:Node,b:Node,liveAdditionSession=false):TableImportXmlDifference|undefined {
  const visible=(node:Node)=>Array.from(node.childNodes).filter(child=>canonical(child,undefined,liveAdditionSession)!==undefined)
  const attrs=(node:Element)=>Array.from(node.attributes).filter(attr=>attr.namespaceURI!==XMLNS && !(liveAdditionSession && liveAdditionSessionAttribute(node,attr)))
    .sort((x,y)=>JSON.stringify([x.namespaceURI,x.localName,x.value]).localeCompare(JSON.stringify([y.namespaceURI,y.localName,y.value])))
  const result=(category:TableImportXmlDifference["category"],path:string):TableImportXmlDifference=>({category,path:bounded(path)})
  const visit=(left:Node,right:Node,path:string):TableImportXmlDifference|undefined=>{
    if (liveAdditionSession?sameLiveTableImportPreimageXml(left,right):sameTableImportXml(left,right)) return
    if (left.nodeType!==right.nodeType) return result("node_kind",path)
    if (left.nodeType===1) {
      const x=left as Element,y=right as Element
      if(x.namespaceURI!==y.namespaceURI||x.localName!==y.localName)return result("element_name",path)
      const xa=attrs(x),ya=attrs(y)
      const key=(attr:Attr)=>JSON.stringify([attr.namespaceURI,attr.localName])
      const oldAttrs=new Map(xa.map(attr=>[key(attr),attr])),newAttrs=new Map(ya.map(attr=>[key(attr),attr]))
      for(const old of xa)if(!newAttrs.has(key(old)))return result("attribute_presence",`${path}/@${safeName(old,true)}`)
      for(const now of ya)if(!oldAttrs.has(key(now)))return result("attribute_presence",`${path}/@${safeName(now,true)}`)
      for(const old of xa) {
        const now=newAttrs.get(key(old))!,at=`${path}/@${safeName(old,true)}`
        if(old.value!==now.value)return result("attribute_value",at)
        if(old.namespaceURI===MC&&old.value.trim().split(/\s+/).some(prefix=>x.lookupNamespaceURI(prefix)!==y.lookupNamespaceURI(prefix)))return result("namespace_binding",at)
      }
      const xc=visible(x),yc=visible(y)
      for(let i=0;i<Math.max(xc.length,yc.length);i++) {
        const old=xc[i],now=yc[i],at=childPath(path,old??now!,i)
        if(!old||!now)return result("node_count",at)
        const diff=visit(old,now,at);if(diff)return diff
      }
    }
    return result(left.nodeType===3||left.nodeType===4?"text":"node_value",path)
  }
  return visit(a,b,a.nodeType===1?`/${safeName(a as Element)}[1]`:"/node[1]")
}
export const firstTableImportXmlDifference = (a:Node,b:Node) => firstDifference(a,b)
export const firstLiveTableImportPreimageXmlDifference = (a:Node,b:Node) => firstDifference(a,b,true)
