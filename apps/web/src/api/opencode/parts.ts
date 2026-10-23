// Shape returned by the search-document tool in its part metadata.citations.
export type Citation = {
  documentName: string
  docPath: string
  section: string
  structuralPath?: string
  excerpt: string
  charStart: number
  charEnd: number
  // search-document's relevance score; a quotation cite verified has none.
  score?: number
}
