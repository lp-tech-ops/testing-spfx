# CAO Case Management System — Search & Report Component
## Technical Context & Design Documentation

> **Purpose of this document:** Provide complete context to an LLM (or developer) tasked with
> building a custom **Search / Report** component *inside* the existing CAO Case Management
> SharePoint Framework (SPFx) web part. It captures the requirement, the platform research we
> validated against PnP Modern Search, the exact problems we hit and how we solved them, the
> reusable KQL/query artifacts, and the recommended build approach.
>
> **Audience:** LLM / SPFx developer.
> **Environment:** SharePoint Online, Microsoft 365, Arlington County (CAO — County Attorney's Office).
> **Dev site:** `https://arlingtonva.sharepoint.com/sites/CAO-CaseManagementSite-Dev`
> **Sanitize note:** Replace tenant/site URLs and internal names with placeholders if this file is
> shared outside the org.

---

## 1. Background & Requirement

### 1.1 The origin
The CAO Case Management System needs a **Search and Reporting** feature modeled on the
**MacroView** "Search and Reporting" panel. Rather than build a full search engine from scratch,
the goal is to reuse proven SharePoint search building blocks and only custom-develop the gaps.

### 1.2 Required capabilities (from the MacroView reference UI)
1. **Keyword search** with four modes: *All of these words*, *The exact phrase*,
   *Any of these words*, *None of these words*.
2. **Filters/refiners**: Result type (file type), "Checked out to me", "My Documents",
   plus **case-specific metadata** (e.g., Assignment Managed By, Doc Set Number, Status).
3. **Site-scoped search** returning a **results grid** (Type, Name, Modified Date, Size, URL).
4. **Order by** (Relevance / Modified).
5. **Pagination** (e.g., 1–50 of N).
6. **Export the results table to Excel**.

### 1.3 Key scope decision
The system stores cases as **Document Sets** using a specific content type,
**"Assignment Document Set"**, spread across **~20 document libraries** on the same site.
The search must target **only Document Sets of this content type**, aggregated across **all**
libraries — not one library at a time.

**Content Type ID (Assignment Document Set):**
```
0x0120D500D301AB8D7B30C54F9368C1B34BB9803E0100912C42DDCEE43A4D9A2C24E2477259E9
```
(Base Document Set content type prefix used in the production flow: `0x0120D520`.)

---

## 2. Approach Evaluation

### 2.1 Options considered
| Option | Summary | Verdict |
|---|---|---|
| Build from scratch | Full custom search UI + engine | Too much effort |
| **PnP Modern Search** web parts | Community/OSS SPFx search web parts | Best base (~80% fit) |
| Custom SPFx web part | Small standalone component reusing SP Search API | Best for full MacroView parity + export |

### 2.2 Why PnP Modern Search
- Free, open-source, actively maintained, deployable standalone.
- Ships four connectable web parts: **Search Box**, **Search Filters**, **Search Verticals**,
  **Search Results**.
- Supports KQL, managed-property refiners, custom Handlebars/web-component layouts.
- Searches the **Microsoft/SharePoint Search index** → respects item-level permissions and
  index freshness. (Behavior differs slightly from MacroView's direct DMS queries — worth
  flagging to stakeholders.)

### 2.3 The one true gap
PnP Modern Search has **no native "export results table to Excel/CSV"** button.
This is a long-standing, still-unimplemented feature request on their GitHub. Export must be
custom-built (see §9).

---

## 3. PnP Modern Search Architecture (validated)

### 3.1 The web parts and data flow
```
Search Box ──(query text)──▶ Search Results ──(exposes refiner values)──▶ Search Filters
                                    ▲                                            │
                                    └──────────(selected filter values)─────────┘
```

- **Search Results** is the *engine*: it holds the **data source**, **query template**
  (site + content-type scope), **Selected properties**, layout, sort, and pagination.
- **Search Filters** is a *dependent*: it does **not** query the site. It derives its filter
  values (refiners/aggregations) from what a connected Search Results web part already retrieved.
  **Therefore it can only connect to a Search Results web part — never to a site or library.**
- **Search Box** feeds the dynamic query text into Search Results.

### 3.2 The two-way connection (critical, commonly missed)
Filters ↔ Results must be configured in **both** web parts or filters silently stay empty:

**A. In Search Results:**
- Property pane → Settings page → enable **Use filters / sort**.
- **Get filters configuration from this Web Part** → select the Search Filters web part.

**B. In Search Filters:**
- Property pane → Settings → **Get filter values from these Web Parts** → select the Search
  Results web part.
- Optional: operator between filters (AND/OR), default OR.

**C. Search Box → Search Results:**
- In Search Results → **Input query text** → switch Static → **Dynamic** → connect to the
  Search Box web part. Set a **default query text** so results (and thus available fields) load
  on page open.

**D. Selected properties (enables filters to have anything to show):**
- Search Results → panel 1 → **Selected properties** → add every managed property you want as a
  column *and/or* as a filter source.
- Search Filters → **Filters settings** → add a filter per property, pick a template.

> **Gotcha:** Property picker dropdowns only list fields **after** Results has actually returned
> data. If a dropdown is empty, **type the managed property name manually and press Enter**
> (this is expected for custom `RefinableStringXX`).

---

## 4. Managed Properties, Crawled Properties & Refiners

### 4.1 The core chain
```
Custom column ──(crawl)──▶ Crawled property (ows_<InternalName>)
             ──(map in Search Schema)──▶ RefinableStringXX / RefinableDateXX
             ──(add to Selected properties)──▶ Search Results
             ──(add as filter)──▶ Search Filters
```

### 4.2 Key rules (SharePoint Online)
- Custom columns **do not** appear as filters automatically. Only a few built-ins
  (Author, FileType, Created, etc.) are pre-mapped.
- SPO **does not** allow creating a *new* managed property flagged "Refinable". You must reuse a
  built-in empty **`RefinableString00–199`** (text) or **`RefinableDateXX`** (dates) slot and map
  your crawled property into it.
- **Crawled property naming:** `ows_<InternalName>`. Spaces become `_x0020_`
  (e.g., "Case Number" → `ows_Case_x0020_Number`). Avoid spaces in new column internal names.
- **Managed metadata columns** produce multiple crawled properties; map `ows_<ColumnName>`
  (NOT the `ows_taxId_...` variant).
- **Person/Group columns** produce multiple crawled properties; map the plain `ows_<InternalName>`
  (NOT `ows_q_USER_...` / ID variants) or the filter will show **numeric IDs instead of names**.
- **RefinableStrings are refinable but NOT full-text searchable.** To let users *type* a value in
  the search box and match it, additionally map the same crawled property to a **separate custom
  managed property** flagged **Queryable/Searchable**.

### 4.3 Propagation
- New mappings are **asynchronous**: typically **15–30 minutes**, but **up to 24 hours**.
- Mapping does **not** retroactively update already-indexed items. Force a reindex:
  **Library Settings → Advanced settings → Reindex Document Library**, or
  **Site Settings → Search and offline availability → Reindex site**.
- The column must be **populated** on items for a crawled property to exist at all.
- **Content Approval** on a library means search only crawls *approved* items (can cause missing
  values).

---

## 5. Document Set–Specific Considerations

1. **Doc Set vs. child documents:** A Document Set is indexed as its own item (folder-like),
   separate from the files inside it. Columns on the Doc Set live on the Doc Set item; child files
   have their own metadata. Decide which you're filtering:
   - Filtering **Document Sets** → scope by the Assignment Document Set content type (our case).
   - Filtering **child documents** → scope by `IsDocument:True`.
2. **Shared Columns** push values to child documents only after re-index; empty/un-crawled →
   no crawled property yet.
3. Optional `contentclass` values for precision:
   - `sts_listitem_documentlibrary` → folders/document sets (list items in a doc library).
   - Filtering by content type ID is usually sufficient and cleaner (below).

---

## 6. Content-Type Filtering — the "20 libraries" solution

Filtering by content type solves the multi-library problem elegantly: you scope
**Search Results** once, and every library using that content type is included automatically
(including libraries added later).

### 6.1 The three built-in content-type managed properties (no mapping needed)
| Property | Contains | Notes |
|---|---|---|
| `ContentTypeId` | Full hex ID `0x0120D500...` | Best: precise + wildcard for child types |
| `SPContentType` | Friendly name "Assignment Document Set" | Readable but brittle to renames |
| `ContentType` | MIME + name combined | Needs `Contains` qualifier |

### 6.2 Query template for Search Results
```
{searchTerms} Path:"https://arlingtonva.sharepoint.com/sites/CAO-CaseManagementSite-Dev" ContentTypeId:0x0120D500D301AB8D7B30C54F9368C1B34BB9803E0100912C42DDCEE43A4D9A2C24E2477259E9*
```

Readable alternative:
```
{searchTerms} Path:"https://arlingtonva.sharepoint.com/sites/CAO-CaseManagementSite-Dev" SPContentType:"Assignment Document Set"
```

**Critical KQL rules learned:**
- **A space after `{searchTerms}` is mandatory.** Without it, a search for "Test" becomes
  `TestPath:https://...` which returns 0 results. (This exact bug is documented in the PnP
  GitHub Document Set discussion.)
- **Trailing `*`** on `ContentTypeId` = prefix match; also catches **child content types**
  derived from Assignment Document Set (future-proofing).
- Content-type scoping alone returns **only the Doc Set items** (child docs excluded), so no
  separate `contentclass` filter is needed.
- Operators (`AND`, `OR`, `NOT`) must be **UPPERCASE**; queries are otherwise case-insensitive.

### 6.3 Keyword modes → KQL (the four MacroView boxes)
- All of these words → `word1 word2` (space = AND)
- The exact phrase → `"word1 word2"`
- Any of these words → `word1 OR word2`
- None of these words → `-word1` or `NOT word1`

---

## 7. Worked Example: mvManagedBy & mvDocSetNumber

### 7.1 "Assignment Managed By" (internal name `mvManagedBy`)
- Crawled property: `ows_mvManagedBy`.
- If it's a **Person/Group** column, map the plain `ows_mvManagedBy` (avoid `ows_q_USER_...`).
- Map to an unused `RefinableStringXX`, set an alias (e.g., `AssignmentManagedBy`).
- Add to Selected properties (Results), then add as a filter.
- **Filter template:** People (nice rendering) or Check box (safest). There is **no "Text"
  template** — text properties use the **Check box** template.

### 7.2 "Doc Set Number" (internal name `mvDocSetNumber`, mapped to `RefinableString135`)
- Crawled property: `ows_mvDocSetNumber`.
- Mapped to **RefinableString135**.
- Because it's a **unique identifier per Document Set**, a Check box refiner is poor UX
  (hundreds of one-off values; refiner values capped at **# of values = 100**, max **1000**).
- **Better:** use the **Combo** template (searchable type-ahead) for browsing, or make it
  **searchable** and drive exact lookups from the Search Box via `mvDocSetNumber:<value>`.

---

## 8. Problems Faced & Solutions (chronological)

| # | Problem | Root cause | Solution |
|---|---|---|---|
| 1 | Search Filters only showed **web parts** (not the site) as "Available connections" | By design — Filters derive values from a Results web part, not a site | Scope the **Results** web part to the site; connect Filters to Results (two-way) |
| 2 | Filters showed nothing / configured but empty | Missing the **return-path** connection and/or Selected properties | Configure connection in **both** web parts; add managed props to Selected properties |
| 3 | Document Set **custom columns** not available as filters | Custom columns never auto-appear; need refinable mapping | Map `ows_<col>` → `RefinableStringXX`; reindex; add to Selected + Filters |
| 4 | Only one library visible but ~20 libraries have the content type | Was thinking per-library | Scope by **ContentTypeId** (content-type filtering covers all libraries) |
| 5 | Mapped `RefinableString135` but not visible in Search Results | Picker doesn't auto-list custom refinables; propagation delay | **Type it manually + Enter**; reindex; wait 15 min–24 hr; verify via **Debug** layout |
| 6 | Filter **Template** field mandatory but had **no "Text" option** | "Template" = display style, not data type | Use **Check box** (text default) or **Combo** (searchable) or **People** |
| 7 | Person filter risk: numbers instead of names | Mapped the ID/`ows_q_USER_` variant | Map the plain `ows_<InternalName>` |
| 8 | Production **Send an HTTP request** search failing during remap | Item indexed under old property while new mapping propagates | Add **OR** condition to query both old and new properties (see §8.1) |
| 9 | Assumed **"Enable download"** = Excel export | It downloads **files/zip**, not the metadata table | Keep for file download; build custom export for metadata (see §9) |

### 8.1 Production flow fix — OR condition (belt & suspenders during reindex)
Flow: **CAO - Intake form - Notification Process flow - PROD** →
action **Send an HTTP request to SharePoint** (inside a `Do until` loop calling
`_api/search/postquery`).

**Original body:**
```json
{
  "request": {
    "Querytext": "mvDocSetNumberOWSTEXT:\"@{triggerOutputs()?['body/DocSetID']}\" ContentTypeId:0x0120D520*",
    "SelectProperties": { "results": [ "Title", "mvDocSetNumberOWSTEXT" ] },
    "RowLimit": 5
  }
}
```

**Corrected body (OR + parentheses + returns which prop matched):**
```json
{
  "request": {
    "Querytext": "(mvDocSetNumberOWSTEXT:\"@{triggerOutputs()?['body/DocSetID']}\" OR RefinableString135:\"@{triggerOutputs()?['body/DocSetID']}\") ContentTypeId:0x0120D520*",
    "SelectProperties": { "results": [ "Title", "mvDocSetNumberOWSTEXT", "RefinableString135" ] },
    "RowLimit": 5
  }
}
```

**Why the parentheses matter:** without grouping, `A:"x" OR B:"x" ContentTypeId:...*` binds the
content-type restriction only to the `B` branch, letting the `A` branch match any content type.
Parentheses force the content-type filter onto **both** branches. Keep the OR through the reindex
window (async, up to 24 hr); simplify later once `RefinableString135` is confirmed populated for
all items. Optionally add `"TrimDuplicates": false` or a `contentclass` guard to avoid matching a
child document that inherits the Doc Set number.

---

## 9. Export Feature — findings (important for the component)

### 9.1 "Enable download" is NOT a metadata export
Native **Enable download** (Details List layout, v4.12+) downloads the **actual selected files**
(single file as-is; multiple/folders as a **.zip**) — identical to SharePoint's library Download.
It does **not** export the results **table/columns**.

Requirements if used: on **Selected properties** include `ContentTypeId`, `NormListID`,
`NormUniqueID`, `SPWebUrl`; layout = Details List; enable **Allow items selection**
(+ Allow multiple selection); enable **Enable download**. Works across libraries within the same
tenant/hostname.

### 9.2 Metadata/table export (the MacroView "Export to Excel") — must be custom
PnP Modern Search maintainers confirmed exporting result field values to CSV/XLSX is **not
built-in** and would require a custom PR/component.

**Constraint that kills the naive approach:** PnP Modern Search sanitizes all Handlebars templates
with **DOMPurify**. It strips `<script>`, `<button>` formaction, and **all `on*` attributes
(onclick, etc.)**. So you **cannot** add an inline `onclick` Export button to a custom layout.

### 9.3 Viable export routes
| Option | What it is | Effort | Fit for CAO |
|---|---|---|---|
| **A. ListView Command Set** (`js-command-selecteditems-export`) | Official PnP sample; "Export to Excel" (.xlsx) command on a **library view**, using the view's columns | Low (ready-made) | Good for per-library browsing, not the cross-library search grid |
| **B. Custom Web Component** (extensibility library) | Registered `<...>` component (survives DOMPurify) with an Export button on the **search page** | Med-High (SPFx dev; must match main solution SPFx version — v4.21.0+ uses **SPFx 1.22.2 + Heft**) | Good if staying on PnP Search UI |
| **C. Standalone SPFx web part** | Own UI: keyword boxes + grid + export; runs `_api/search/postquery` | Med | **Best for full MacroView parity + export** |
| **D. PnP PowerShell** (`Submit-PnPSearchQuery … \| Export-Csv`) | Headless CSV export | Low | Great for scheduled/admin reporting |

### 9.4 Excel generation techniques
- **`xlsx` (SheetJS) + `file-saver`** — simple/quick, plain grid dump.
- **ExcelJS** — formatting, styling, headers, multiple sheets (better for polished CAO reports).
- **ListView Command Set** (Option A) already wraps xlsx.

---

## 10. Recommended Build — Component INSIDE the existing Case Management web part

> The user's decision: build the search/report as a **component within the existing CAO Case
> Management SPFx web part** (not a separate PnP page). This means DOMPurify/Handlebars
> constraints do **not** apply — you have full React/SPFx control. Reuse PnP Search *concepts and
> KQL*, not the PnP web parts themselves.

### 10.1 Architecture
- A React component (e.g., `AssignmentSearchReport`) mounted inside the existing web part.
- **Search service** using PnPjs: `spfi().using(SPFx(this.context))` → `search(...)`, OR call
  `_api/search/postquery` directly (mirrors the proven PROD flow query).
- **State:** query text (four keyword modes), selected filters, sort, paging, results, loading.
- **Export:** a real React `onClick` button → ExcelJS/xlsx → download (no DOMPurify limits here).

### 10.2 Reusable artifacts (lift directly)
- **Base scope / query template:**
  ```
  <keywords built from 4 modes> ContentTypeId:0x0120D500D301AB8D7B30C54F9368C1B34BB9803E0100912C42DDCEE43A4D9A2C24E2477259E9* Path:"https://arlingtonva.sharepoint.com/sites/CAO-CaseManagementSite-Dev"
  ```
- **Keyword-mode → KQL builder** (see §6.3).
- **Managed properties to select/return** (extend as needed):
  `Title, Path, FileType, Size, LastModifiedTime, Author, Editor,
   SPWebUrl, ContentTypeId, NormListID, NormUniqueID,
   mvDocSetNumberOWSTEXT, RefinableString135 (Doc Set Number),
   RefinableStringXX (Assignment Managed By / mvManagedBy)`
- **Filter fields** (refinable managed props): map each Doc Set column
  `ows_<internal>` → `RefinableStringXX` first (§4).
- **OR-during-remap pattern** for any property mid-migration (§8.1).

### 10.3 Managed-property mapping checklist (do before coding filters)
For each Doc Set column to be filterable/reportable:
1. Confirm `ows_<InternalName>` exists (Search Schema → Crawled Properties).
2. Map to unused `RefinableStringXX` (text) / `RefinableDateXX` (date); set alias.
3. (If searchable-by-typing needed) also map to a new Queryable custom managed property.
4. Reindex library/site; wait for propagation.
5. Verify values via Debug (PnP) or `Submit-PnPSearchQuery`.

**Known field mappings so far:**
| Column (internal) | Type | Crawled prop | Managed prop |
|---|---|---|---|
| `mvDocSetNumber` | Text | `ows_mvDocSetNumber` | `RefinableString135` (+ legacy `mvDocSetNumberOWSTEXT`) |
| `mvManagedBy` | Person/Group (confirm) | `ows_mvManagedBy` | `RefinableStringXX` (map plain variant) |

### 10.4 UI parity checklist (MacroView)
- [ ] Four keyword inputs (All / Exact / Any / None) → KQL builder.
- [ ] Result type (FileType) filter.
- [ ] "Managed By" filter (RefinableString, People/Combo).
- [ ] Doc Set Number search (searchable prop or Combo).
- [ ] Results grid: Type, Name, Modified, Size, URL (+ Doc Set Number, Managed By).
- [ ] Order by (Relevance / Modified).
- [ ] Pagination (page size 50).
- [ ] **Export to Excel** button (ExcelJS/xlsx) — exports current filtered result set.
- [ ] (Optional) File download of selected Doc Sets (mirror "Enable download").

### 10.5 Gotchas to bake into the component
- Always put a **space** between keyword text and property restrictions in KQL.
- Uppercase `AND/OR/NOT`.
- Search reflects the **index** (permission-trimmed, may lag content by minutes).
- Handle **RowLimit / paging** (search caps; use paging tokens for large sets).
- Refiner value caps (100 default) don't apply to your own grid, but do to any PnP-style refiner.
- Person columns: display name vs. ID variant.
- Content Approval libraries: only approved items crawled.

---

## 11. Key References (for the LLM to consult)
- PnP Modern Search — Search Results / Filters usage, connections, layouts, templating,
  extensibility, custom web components (microsoft-search.github.io/pnp-modern-search).
- PnP "Enable download" feature write-up (Patrik Hellgren).
- Content-type filtering with search-driven UIs (Joanne C. Klein; Marc D. Anderson).
- Crawled property naming conventions (SharePoint Maven).
- KQL syntax reference (Microsoft Learn).
- Export samples: `pnp/sp-dev-fx-extensions` → `js-command-selecteditems-export`;
  SPFx xlsx + FileSaver + ExcelJS methods.
- PnP PowerShell `Submit-PnPSearchQuery … | Export-Csv` script sample.

---

## 12. TL;DR for the implementing LLM
Build a React sub-component in the existing CAO Case Management SPFx web part that:
1. Constructs KQL from four keyword modes + filters, always scoped by
   `ContentTypeId:0x0120D500...*` and the CAO-Dev site path (with a leading space rule).
2. Queries SharePoint Search (PnPjs `search()` or `_api/search/postquery`) returning the managed
   properties listed in §10.2 (map any missing Doc Set columns to `RefinableStringXX` first).
3. Renders a paginated, sortable results grid mirroring MacroView.
4. Provides a native React **Export to Excel** button (ExcelJS/xlsx) over the current filtered
   results — this is the piece PnP Modern Search cannot do natively.
5. (Optional) offers file download of selected Doc Sets like PnP's "Enable download".
Because this lives inside your own web part, you are free of PnP's DOMPurify template restrictions.
