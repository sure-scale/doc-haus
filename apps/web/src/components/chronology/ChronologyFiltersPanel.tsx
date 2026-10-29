import JurisdictionSelect from "../JurisdictionSelect"
import DateInput from "./DateInput"
import type { DateInputFormat } from "./date-input"
import type { ChronologyFilters } from "./model"
import "./chronology-filters.css"

export default function ChronologyFiltersPanel({
  filters,
  onChange,
  documentNames,
  people,
  dateFormat,
}: {
  filters: ChronologyFilters
  onChange: (filters: ChronologyFilters) => void
  documentNames: string[]
  people: string[]
  dateFormat: DateInputFormat
}) {
  const update = (patch: Partial<ChronologyFilters>) => onChange({ ...filters, ...patch })
  return (
    <div
      id="chronology-filters"
      className="chronology-filters chronology-filters-panel"
      role="group"
      aria-label="Event filters"
    >
      <div className="chronology-filter-primary">
        <div className="chronology-filter-field">
          <span className="chronology-filter-label">Documents</span>
          <JurisdictionSelect
            jurisdictions={documentNames.map((name) => ({ code: name, name }))}
            selected={filters.document ? [filters.document] : []}
            onChange={(codes) => update({ document: codes[0] ?? "" })}
            multiple={false}
            label="Document"
            searchLabel="Search documents"
            emptyLabel="All documents"
            showCode={false}
            allowClear
          />
        </div>
        <div className="chronology-filter-field">
          <span className="chronology-filter-label">People</span>
          <JurisdictionSelect
            jurisdictions={people.map((name) => ({ code: name, name }))}
            selected={filters.person ? [filters.person] : []}
            onChange={(codes) => update({ person: codes[0] ?? "" })}
            multiple={false}
            label="Person / organisation"
            searchLabel="Search people and organisations"
            emptyLabel="All people"
            showCode={false}
            allowClear
          />
        </div>
        <fieldset className="chronology-filter-date-range">
          <legend>Date range</legend>
          <div className="chronology-filter-dates">
            <div className="chronology-filter-date-bound">
              <span className="chronology-filter-date-prefix" aria-hidden="true">
                From
              </span>
              <DateInput
                label="From date"
                value={filters.from}
                format={dateFormat}
                onChange={(value) => update({ from: value ?? "" })}
              />
            </div>
            <div className="chronology-filter-date-bound">
              <span className="chronology-filter-date-prefix" aria-hidden="true">
                To
              </span>
              <DateInput
                label="To date"
                value={filters.to}
                format={dateFormat}
                onChange={(value) => update({ to: value ?? "" })}
              />
            </div>
          </div>
        </fieldset>
      </div>
      <div className="chronology-filter-secondary">
        <label className="chronology-filter-field">
          <span>Review</span>
          <select value={filters.review} onChange={(change) => update({ review: change.target.value })}>
            <option value="">All review states</option>
            <option value="new">New, not opened</option>
            <option value="unreviewed">Unreviewed</option>
            <option value="reviewed">Reviewed</option>
            <option value="attention">Needs attention</option>
          </select>
        </label>
        <label className="chronology-filter-field">
          <span>Significance</span>
          <select value={filters.significance} onChange={(change) => update({ significance: change.target.value })}>
            <option value="">All significance</option>
            <option value="key">Key (AI assessment)</option>
            <option value="relevant">Relevant</option>
            <option value="background">Background</option>
          </select>
        </label>
        <label className="chronology-filter-field">
          <span>Visibility</span>
          <select
            value={filters.hidden}
            onChange={(change) => update({ hidden: change.target.value as ChronologyFilters["hidden"] })}
          >
            <option value="exclude">Visible events</option>
            <option value="include">Include hidden</option>
            <option value="only">Hidden only</option>
          </select>
        </label>
      </div>
      <fieldset className="chronology-filter-additional">
        <legend>Additional filters</legend>
        <div className="chronology-filter-flags">
          <label className="chronology-check">
            <input
              type="checkbox"
              checked={filters.undated}
              onChange={(change) => update({ undated: change.target.checked })}
            />
            Undated
          </label>
          <label className="chronology-check">
            <input
              type="checkbox"
              aria-label="Lawyer-starred key events"
              checked={filters.key}
              onChange={(change) => update({ key: change.target.checked })}
            />
            Lawyer-starred
          </label>
          <label className="chronology-check">
            <input
              type="checkbox"
              checked={filters.conflict}
              onChange={(change) => update({ conflict: change.target.checked })}
            />
            Conflicts / disputes
          </label>
        </div>
      </fieldset>
    </div>
  )
}
