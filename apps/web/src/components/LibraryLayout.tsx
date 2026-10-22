import { useId, type ReactNode } from "react"
import PageBar from "./PageBar"

// Shared presentation for the firm libraries. Each page owns its data and actions.
export default function LibraryLayout({
  title,
  description,
  libraryTitle,
  count,
  active,
  upload,
  assistant,
  children,
}: {
  title: string
  description: string
  libraryTitle: string
  count?: number
  active: boolean
  upload?: ReactNode
  assistant: ReactNode
  children: ReactNode
}) {
  const heading = useId()
  return (
    <>
      <PageBar title={title} description={description} />
      <div className={`library-layout${active ? " is-active" : ""}${assistant ? "" : " library-layout-single"}`}>
        <section className="library-collection" aria-labelledby={heading}>
          <div className="library-heading">
            <h2 id={heading}>{libraryTitle}</h2>
            {count !== undefined && <span className="library-count" aria-label={`${count} ${libraryTitle.toLowerCase()}`}>{count}</span>}
          </div>
          {upload && <div className="library-upload">{upload}</div>}
          <div className="library-items">{children}</div>
        </section>
        {assistant && (
          <section className="library-builder" aria-label={`${title} assistant`}>
            {assistant}
          </section>
        )}
      </div>
    </>
  )
}
