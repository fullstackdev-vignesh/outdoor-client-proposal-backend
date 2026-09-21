# PPT Template System — Full Business Flow

This file documents the **entire PPT proposal generation flow** — frontend UI, backend
controllers/models, the PPTX generation engine, and every template's individual business logic —
so that any future change/modification request can be understood end-to-end without re-reading
the whole codebase.

Two independent repos are involved:
- Backend: `backend/outdoor-client-proposal-backend`
- Frontend: `frontend/outdoor-client-proposal-frontend`

---

## 1. High-level flow (end to end)

```
Frontend: Create Proposal wizard (5 steps)
   Step 1: Customer Details  (pick Client/Agency)
   Step 2: Site Details      (pick sites — only mediaStatus="available")
   Step 3: PPT Template      (pick one PPTTemplate doc)
   Step 4: Excel Template    (pick one ExcelTemplate doc)
   Step 5: Preview           (review, then "Generate Proposal")
        |
        v
POST /api/proposals  -> creates a Proposal doc (status: "draft")
        |
        v
Proposal Details page: "Without Location" / "With Location" buttons
        |
        v
POST /api/proposals/:id/generate-ppt  { locationMode: "with" | "without" }
        |
        v
proposalController.generatePpt()
   -> populates client + sites (+ site.siteInfoId) + pptTemplate
   -> calls generateProposalPpt(proposal, { locationMode })
        |
        v
proposalFileGenerator.js: generateProposalPpt()
   -> loads the PPTTemplate's underlying .pptx (or master.pptx fallback)
   -> looks at pptTemplate.name to pick ONE of 7 hardcoded template branches
      (adinn-new-template / adinn-photos-only / adinn-customized-format /
       Adinn-Direct-Client-format / publicis-ooh-template /
       Jagran-template-one / Jagran-template-two)
   -> each branch clones/mutates slides via pptxTemplateEngine.js (raw OOXML edits)
   -> uploads finished .pptx to cloud/local storage
   -> saves proposal.generatedPptUrl
        |
        v
Frontend: "Download PPT" link becomes available
```

Excel generation is a **separate, parallel** flow (`generateProposalExcel` /
`excelTemplateEngine.js`) — not covered in depth here since with/without-location does not apply
to Excel.

---

## 2. Frontend flow

### 2.1 Create Proposal wizard — `src/app/(app)/proposals/new/page.tsx`

5-step wizard (`STEPS` array): Customer Details → Site Details → PPT Template → Excel Template →
Preview.

- **Step 1 (Customer)**: toggles `client`/`agency` customer type, searches `/clients`, picks one.
- **Step 2 (Sites)**: infinite-scroll list from `/sites` (filters: search/state/city/status/owner).
  Only sites with `mediaStatus === 'available'` are selectable (others shown disabled) — this is
  also enforced server-side in `createProposal`.
- **Step 3 (PPT Template)**: grid of active `PPTTemplate` docs from `/ppt-templates?status=active`.
- **Step 4 (Excel Template)**: same pattern against `/excel-templates`.
- **Step 5 (Preview)**: summary + PPT/Excel structure preview (static descriptive text, not a real
  render) + "Generate Proposal" button.
- **Generate Proposal** → `POST /proposals` with `{ client, sites, pptTemplate, excelTemplate,
  totalAmount, gstAmount, monthlyAmount }` → creates the `Proposal` doc (`status: 'draft'`).
- From the Preview step, "Download PPT"/"Download Excel" call `POST
  /proposals/:id/generate-ppt|excel` directly (no `locationMode` sent here — defaults to `'with'`
  server-side) and open the resulting file in a new tab.

### 2.2 Proposal Details page — `src/app/(app)/proposals/[id]/page.tsx`

Shown after a proposal is created (`/proposals/[id]`). Key section — **Generate & Download**:

- **"Without Location" button** → `generatePpt('without')` → `POST
  /proposals/:id/generate-ppt` with `{ locationMode: 'without' }`.
- **"With Location" button** → `generatePpt('with')` → same endpoint with `{ locationMode: 'with'
  }`.
- Both re-`GET /proposals/:id` on success to refresh `generatedPptUrl`, and a "Download PPT" link
  appears once that URL exists. Re-clicking either button **regenerates and overwrites**
  `generatedPptUrl` (no separate "with" vs "without" file is kept — only the latest generation
  survives).
- "Generate Excel / Refresh" is a separate, unrelated button (no location mode).

### 2.3 Site Info management (master data for the "Site Information" card)

- `src/components/sites/SiteFormModal.tsx` — the Site Add/Edit form has an optional "Site
  Information" dropdown (loads `GET /site-info`) plus a "+ Add Site Information" quick-create
  modal (`SiteInfoQuickAddModal`, same file) that `POST`s to `/site-info` and auto-selects the new
  entry. Entirely optional — leaving it blank does not affect normal site creation.
- `src/components/sites/SiteInfoManager.tsx` + `src/app/(app)/site-info/page.tsx` — full CRUD
  page (list/create/edit/delete), same pattern as `TemplateManager.tsx`. Linked from the sidebar
  nav (`src/lib/nav.ts`, admin/tl roles only).
- This data only affects **`Adinn-Direct-Client-format`**'s "Site Information" description card
  (see §4.4) — every other template ignores `site.siteInfoId` entirely.

---

## 3. Backend flow

### 3.1 Models

- **`Proposal`** — `client`, `sites[]`, `pptTemplate`, `excelTemplate`, `variant`, amounts,
  `status` (`draft → generated/completed`), `generatedPptUrl`, `generatedExcelUrl`.
- **`Site`** — master fields (specs/location/pricing) + inventory fields (status/bookings). The
  field relevant to PPT generation: `siteInfoId` (optional ref to `SiteInfo`).
- **`SiteInfo`** — `{ title, description }`, master data reused across multiple sites.
- **`PPTTemplate`** — `{ name, description, status, fileUrl/filePath }`. The `name` field is what
  `proposalFileGenerator.js` string-matches to select a template branch (see §3.3).

### 3.2 Controller — `controllers/proposalController.js`

- `createProposal` — validates client + all sites exist and are `mediaStatus: 'available'`.
- `generatePpt` — populates `client`, `sites` (with nested `siteInfoId`), `pptTemplate`; reads
  `req.body.locationMode` (`'without'` only if explicitly sent, else defaults to `'with'` — fully
  backward compatible with any caller that omits it); calls `generateProposalPpt(proposal, {
  locationMode })`; saves `generatedPptUrl` and bumps `status`.
- `generateExcel` — separate, no `locationMode`.

### 3.3 Generation engine — `utils/proposalFileGenerator.js`

`generateProposalPpt(proposal, { locationMode = 'with' } = {})`:

1. Resolves `withLocation = locationMode !== 'without'`.
2. Loads the proposal's `pptTemplate` file via `PptxTemplate.load(...)` (falls back to
   `assets/proposal-templates/master.pptx` if none set).
3. Reads `templateName = pptTemplate?.name` and checks it against 7 hardcoded flags
   (`isAdinnNewTemplate`, `isAdinnPhotosOnly`, `isAdinnCustomizedFormat`,
   `isAdinnDirectClientFormat`, `isPublicisOohTemplate`, `isJagranTemplateOne`,
   `isJagranTemplateTwo`). **Exactly one** `if` branch runs based on this name match; if none
   match, a generic/legacy fallback path runs (kept for old templates, not documented in depth
   here — do not delete it).
4. Sets cover-slide fields (`setCoverFields`/`setCoverDateLabel`) once, before branching.
5. Inside the matched branch: for each selected site (grouped by state/city where the template
   needs divider slides), clones the relevant reference slide(s), swaps in the site's text/image
   data, and — depending on `withLocation` — adds/removes a map box or resizes boxes (see §4 for
   the exact per-template rules).
6. Assembles the final slide order via `tpl.setFinalSlideOrder([...])`.
7. Saves the `.pptx` buffer, uploads it via `uploadFileToCloud` (cloud Spaces, or local `uploads/`
   fallback), filename built by the shared helper below.
8. Returns the uploaded file's URL — caller stores it on `proposal.generatedPptUrl`.

**Filename convention** (all templates, via `buildGeneratedPptFileName(proposal, client, now)`):
```
<ClientName>-<DD>-<Month>-<YYYY>-<proposalId>.pptx
e.g.  ROTN-19-September-2026-PR-MU83CUXH.pptx
```

### 3.4 PPTX engine — `utils/pptxTemplateEngine.js`

No PowerPoint library is used — a `.pptx` is a zip of XML parts, and this class
(`PptxTemplate`, backed by JSZip) edits those XML strings directly (regex-based text/shape
matching), then re-zips. Never uses `PptxTemplate` outside this file's controlled method set.

Key methods (grouped by purpose):

| Method | Purpose |
|---|---|
| `cloneSlide(base, mutations)` | Generic clone: exact text replacements + simple image relId swap. Used where no box-resize/crop is needed (e.g. city/state divider slides, Jagran-template-two's site slide). |
| `cloneAdinnSiteSlide(base, opts)` | Clone with cover-fit image crop (`boxWidthEMU`/`boxHeightEMU`), optional `removeGroupNames`, `removeShapesAtOffset`, `clearImageRelId` + `placeholderText`. Used by adinn-new-template, adinn-customized-format, Adinn-Direct-Client-format. |
| `clonePhotoOnlySlide` / `clonePhotoWithMapSlide` | adinn-photos-only's single-photo vs photo+map site slide. |
| `cloneCaptionPhotoSlide` | Photo + plain-textbox caption (publicis-ooh-template, Jagran-template-one). |
| `resizeTextBox(slidePath, {offX,offY,newWidthEMU})` | Widen/reposition any shape matched by its **exact** original `<a:off>` — generic enough to resize a `<p:pic>`'s own xfrm too. |
| `resizeNamedGroup(slidePath, groupName, {...})` | Resize a named top-level `<p:grpSp>`'s own xfrm (photo-box groups). |
| `insertImageOrPlaceholder(slidePath, relsPath, {...})` | Add a brand-new `<p:pic>` (real map) or a text placeholder ("Insert your map image here") at an arbitrary position — used everywhere a template has **no** built-in map box and one must be added as a new design element. |
| `insertWhiteCover(slidePath, {...})` | Draw a plain white rectangle over a region — used when label/border graphics are baked into the **slide layout** (not the slide itself) and must be visually hidden (Jagran-template-two). |
| `removeShapesAtOffsets(slidePath, offsets)` | Standalone shape-removal pass (by exact `<a:off y>`/`<a:ext cy>`), for clone methods (like `cloneSlide`) that don't support removal inline. |
| `clearBackgroundImage` | Turn a full-slide `<a:blipFill>` into plain white (adinn-customized-format's blank slides 1-2). |
| `setCoverFields` / `setCoverDateLabel` / `setCoverCustomerNameLiteral` | Cover-slide client name/date substitution (keyword-based or literal-text based, per template's demo text). |
| `mirrorRibbonGroup` | adinn-new-template only — mirrors a decorative ribbon shape to the opposite corner. |
| `setFinalSlideOrder` / `removeFromSlideOrder` | Assemble/trim the final slide sequence. |

Route-map images come from `utils/mapService.js#getRouteMapBuffer({fromLat,fromLng,toLat,toLng})`
(free OSRM + staticmap APIs) — only called when **both** the client and the site have
latitude/longitude; otherwise the map box shows the text placeholder instead.

---

## 4. Per-template business logic (all 7 templates)

For every template: reference slide roles, how sites/cities/states are grouped, and the exact
**With Location vs Without Location** behavior.

### 4.1 `adinn-new-template`

- Slide1 cover, Slide2/3 kept as-is (About Us ribbon widened once).
- One clone per site, in selection order:
  - **With Location** → clones **slide5** only (photo + map). Title = `location + size`. Photo in
    bordered box (`rId9`), map box cleared to placeholder or filled by real route map.
  - **Without Location** → clones **slide4** only (photo + Media Specifications panel: City,
    Size, Media type, Illumination, Unit). Title = `location + size` (same as with-location — only
    the map/spec choice differs, not the title).
- Exactly **one** slide role per site per mode (not both, unlike the original reference which had
  both slide4 and slide5 for every site).
- This template is the **reference convention** every other template's with/without logic is
  matched against (title format, map placeholder style, etc.).

### 4.2 `adinn-photos-only`

- One city-divider + one site-photo slide per site (grouped by city).
- Caption always shows location + size, both modes.
- **With Location** → `clonePhotoWithMapSlide`: photo box narrowed (5,674,400 EMU), real/placeholder
  map box added beside it (3,674,401 EMU, 100,000 EMU gap).
- **Without Location** → `clonePhotoOnlySlide`: full-width single photo, no map.

### 4.3 `adinn-customized-format`

- Slide1/2 forced blank (`clearBackgroundImage`) — meant for the user to fill in manually later.
- Slide3 = site-detail template (two image boxes: main photo `rId2`, smaller `rId3`), cloned once
  per site. Slide5 = "Thank You" (kept verbatim, always last).
- Title always `location – size`, both modes.
- **With Location** → `rId3` box shows a real route map (or "Insert your map image here"
  placeholder) instead of a duplicate site photo; photo box width 10,576,310.
- **Without Location** → `Group 2` (photo) widened to 16,766,194 (fills the whole area); `Group 4`
  (map box) **and** its two decorative pin-icon groups (`Group 18`, `Group 22` — sit inside the
  same x-range as the map box) are all removed so nothing floats on top of the widened photo.

### 4.4 `Adinn-Direct-Client-format`

- Same visual family as adinn-new-template (cover/about/why-us kept as-is), but its per-site slide
  has **no Illumination/Unit fields**, and instead an optional **"Site Information" card**
  (gradient card + description) sourced from `site.siteInfoId.description` (the SiteInfo master
  data, §2.3). If a site has no linked SiteInfo, the whole card group (`Group 29`) + its background
  shape are removed — never shown empty.
- One clone per site from slide4; slides 5/6 in the reference file are unused extra demo examples
  (excluded from `remainingSlides`), slide7 "Thank You" is the real closer.
- **Without Location** → title = location only (size shown separately in the visible Media Spec
  panel: City/Size/Media type). Full Media Specifications panel + Site Info card (if present) both
  shown. Photo box uses `DIRECT_CLIENT_BOX_WIDTHS`.
- **With Location** → title = `location + size` (matches adinn-new-template exactly). Photo box
  narrowed to 10,484,172 EMU (adinn-new-template's own slide5 photo width) and a map box
  (offX 11,605,227, width 6,508,010 — also adinn-new-template's own values) is inserted, ending
  near the slide's right edge with only a small margin. The **entire** Media Specifications panel
  (heading + City/Size/Media type rows) and the Site Info card are hidden — only photo + map +
  title remain.

### 4.5 `publicis-ooh-template`

- Reference slide1 (cover) kept verbatim. Slide2 = state-divider, slide3 = city-divider, slide4 =
  site-detail. Sites grouped by state → city (first-seen order).
- Caption always shows location + size, both modes.
- **With Location** → photo box narrowed 9,772,650 → 5,872,650 EMU; map box (3,800,000 EMU, 100,000
  gap) inserted beside it (real map or placeholder).
- **Without Location** → photo box stays full width (9,772,650), no map.

### 4.6 `Jagran-template-one`

- Reference slide1 = city-divider, slide2 = site-detail. Sites grouped by city.
- Caption always shows location + size, both modes.
- **With Location** → photo box narrowed 8,072,494 → 4,772,494 EMU; map box (3,200,000 EMU, 100,000
  gap) inserted beside it.
- **Without Location** → photo box stays full width (8,072,494), no map.
- Slides 3-6 in the reference file are unused extra demo content — never referenced in the final
  order (no distinct "Thank You" slide exists in this template's reference file).

### 4.7 `Jagran-template-two`

- Reference slide1 ("Jagran Engage" title) and the last slide ("Thank you") kept verbatim. Slide2 =
  city-divider (red text preserved via `preserveRedShapes`). Slide4 = site-detail (used instead of
  slide3 because its Width/Height reference values are distinct, making exact-text replacement
  unambiguous). Sites grouped by city.
- Site-detail slide layout: full-bleed site photo on the left (~60% of slide width), and a Media
  Specifications panel on the right (Location/State/City/Media/Type/Size/W/H/Duration) — **note:**
  the "State:"/"City:"/etc. **labels and borders are baked into the slide layout**
  (`slideLayout1.xml`), not the slide itself, so simply removing the slide's own value shapes is
  not enough to hide them.
- **With Location** → title = `location + width x height`. All Media Spec value shapes removed,
  **plus** a plain white rectangle (`insertWhiteCover`) drawn over the whole panel area (below the
  title, which stays visible) to blank out the layout's baked-in labels/borders, then a real
  route map (or placeholder) is inserted on top of that cover.
- **Without Location** → title = location only; full Media Specifications panel shown exactly as
  in the reference file; no cover, no map.

---

## 5. Quick reference — "which button changes what"

| Template | Title changes with mode? | What "With Location" adds/shows | What "Without Location" shows instead |
|---|---|---|---|
| adinn-new-template | No (always location+size) | slide5 (photo+map) | slide4 (photo+specs) |
| adinn-photos-only | No (always location+size) | narrow photo + map box | full-width photo |
| adinn-customized-format | No (always location–size) | real map / placeholder in 2nd box | photo widened, map+pins removed |
| Adinn-Direct-Client-format | Yes (+size only with-location) | narrow photo + map, specs+SiteInfo card hidden | full specs panel + Site Info card, full photo |
| publicis-ooh-template | No (always location+size) | narrow photo + map box | full-width photo |
| Jagran-template-one | No (always location+size) | narrow photo + map box | full-width photo |
| Jagran-template-two | Yes (+size only with-location) | spec panel hidden (white cover) + map | full spec panel, no map |

---

## 6. Extending this system (how to add a new template's mode logic)

1. Add the template's `templateName` check near the top of `generateProposalPpt` (`isX =
   templateName === '...'`).
2. Inspect the reference `.pptx`'s relevant slide XML directly (via a throwaway Node script using
   `PptxTemplate.load()` + `readText()`) to find exact `<a:off>`/`<a:ext>` values, shape names, and
   text-run content — **never guess EMU values**.
3. Reuse an existing clone method (`cloneSlide`/`cloneAdinnSiteSlide`/`cloneCaptionPhotoSlide`) if
   the slide's structure matches; otherwise reuse the generic post-clone helpers
   (`resizeTextBox`, `resizeNamedGroup`, `insertImageOrPlaceholder`, `insertWhiteCover`,
   `removeShapesAtOffsets`) rather than writing new one-off XML logic.
4. Always verify with a standalone Node script against the **real reference file** before wiring
   into the generator — check for `NaN`, confirm shape counts/removals, and confirm `tpl.save()`
   succeeds — this project has no automated test suite for PPT generation.
5. Regression-check the *other* mode and any onother template that shares a helper method, since
   several methods (`resizeTextBox`, `insertImageOrPlaceholder`, etc.) are shared across templates.
