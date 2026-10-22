// Compact top bar shared by the index pages (Matters and the firm libraries),
// matching the matter page's head so the bar does not jump between them.
export default function PageBar({ title, description }: { title: string; description: string }) {
  return (
    <header className="page-head page-bar">
      <h1>{title}</h1>
      <p className="page-bar-sub">{description}</p>
    </header>
  )
}
