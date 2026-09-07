# Changelog — DS Styles Extractor

## v1.8.1 — 4 September 2026

### New: a name of your own for each collection
- **Per collection** in the Naming bar opens a row for every selected collection, where you can give it a
  different name in the export: `primitives` can go out as `prim`, `secondary` as `sec`, and so on. Leave a
  row blank to keep the real name.
- Each row shows a live example of a real token from that collection, so you can see the result as you type.
- Aliases follow the renamed collections, so a reference to a renamed collection stays correct.
- Like the rest of the naming settings, the renames are remembered after closing the plugin and after
  restarting Figma. They are stored per computer, not in the file.


## v1.8.0 — 4 September 2026

### New: you choose how exported token names are built
Variable names in this design system already carry their own prefix, so adding the collection on top
produced names such as `--primitives-global-global-primitive-type-font-size-xl`. A **Naming** bar now sits
above the format tabs whenever CSS or W3C DTCG is selected on the Variables tab:
- **Collection prefix** — on or off. With it off, the example above becomes
  `--global-primitive-type-font-size-xl`, which is what most of these names want.
- **Merge repeated words** — collapses a word that immediately repeats itself, so `global-global` becomes
  `global`, useful when the collection and the token do overlap.
- **Prefix** — an optional word of your own, for example `ds`, applied to every token.
- A live example on the right shows exactly what the next token name will look like, and the preview and
  the export always agree. Aliases follow the same rules, so references never break.
- The choices are remembered between sessions.
- The raw Figma JSON is deliberately left untouched: it is a faithful dump of the file, used for diffs
  when publishing. W3C DTCG is the one to pick for clean names.


## v1.7.1 — 4 September 2026

### Fixed: the plugin took minutes to open on a large design system
Three separate causes, all in the start-up path.
- **The panel waited for the whole global library.** Nothing was sent to the interface until every global
  variable had been imported and resolved, so on a big file you stared at an empty panel for minutes. The
  local styles and variables are now sent immediately and the panel is usable at once; the global library
  arrives as a second update, with the header saying it is still loading.
- **Global values were resolved one at a time.** Each variable, for each mode, waited for the previous one
  to finish. A few thousand variables across four modes therefore meant thousands of sequential waits. They
  are now resolved in parallel batches.
- **Alias chains asked the API for variables already in memory.** Everything just imported is now cached
  first, so following an alias no longer costs a round trip.

### Fixed: the CSS export was rebuilding its lookup table for every alias
- Introduced in v1.7.0. On a file with 3,150 variables the CSS export took over two seconds; it now takes
  about twenty milliseconds. The table is built once per export.


## v1.7.0 — 4 September 2026

### Fixed: CSS variables from two collections with the same name overwrote each other
- This file has a local **primitives** collection and a global one from the brand library. Flutter and
  W3C DTCG kept them apart, CSS did not, so both produced `--primitives-…` names and the global silently
  won. The global collection now gets a `-global` suffix, and only when a name is genuinely duplicated,
  so token names that were never ambiguous do not move.
- CSS aliases now reference the full path including the collection, for the same reason.

### Fixed: the W3C DTCG export was not usable by a token pipeline
- **References** were written the Figma way, `{primitives/color/background}`, and without the collection.
  The format wants a dot-separated path, so they never resolved. They are now `{primitives.color.background}`.
- **`$type` said `alias`**, which is not a type in the format. It now carries what the alias resolves to.
- **Modes were nested inside the token path**, which made "Light" part of every token's name. The mode is
  now the top level, so each mode is a complete, self-contained token set: references resolve within it,
  and splitting the file into light.json and dark.json is just taking one key. Single-mode collections are
  unchanged.

### Fixed: the pull request title carried a made-up ticket number
- New pull requests were pre-filled with `AJBY00:`, a placeholder that shipped with every export unless
  someone noticed. The title now starts clean, and the hint shows the expected shape instead.


## v1.6.8 — 3 September 2026

### A component that could not be checked now says why
- The note used to claim the API phase had run out of time, whatever the real cause, because that sentence
  was fixed in the code. It now records the actual reason per component and shows it beside the name:
  the token was rejected, Figma asked us to slow down, the request never completed, the API answered with
  a particular status, or it genuinely was not reached in time.
- The list also points out that components from third-party or community kits, such as Material, Android
  and iOS, commonly land here and are harmless, since they are not part of your design system.
- Copying the list now includes the reasons.


## v1.6.7 — 3 September 2026

### The private-components note now explains itself
- Skipping them is stated in full, in the panel and in the canvas report: what a private component is, why
  Figma never publishes it, why the library API cannot say anything about it, and explicitly that these
  **count as correctly linked and do not affect the health score**. A Show them button lists them.

## v1.6.6 — 3 September 2026

### The API phase stops wasting its time on private components
- Nearly everything the scan was running out of time on turned out to be **private components** — the ones
  whose name starts with a dot or an underscore, which Figma never publishes. They are the internal parts
  of published components, so asking the library API about them was guaranteed to come back empty. They
  are now skipped, and the header says how many were skipped and why.
- On a file like the Decibel library this removes the large majority of the lookups, so the phase finishes
  comfortably instead of hitting the clock. Lookups also run twelve at a time now.


## v1.6.5 — 3 September 2026

### The API phase now finishes, and both notes show their components
- Lookups run eight at a time instead of four, and the phase has three minutes instead of one, so the
  several hundred components that were previously abandoned when the clock ran out are now all checked.
  The progress line shows how many are done and how much time is left.
- **Show them** opens a proper list of the components a note is talking about, numbered and scrollable,
  with a copy button inside it, instead of silently copying to the clipboard.
- The "not checked" note now has a Show them button too, so it is always clear which components are meant.


## v1.6.4 — 3 September 2026

### The API phase no longer hangs, and a refused component is explained properly
- Library lookups now run four at a time instead of one by one, with a 60-second budget for the whole
  phase. Whatever is left when the budget runs out is reported as not checked, so a slow or unhelpful API
  can never leave the scan sitting at 80% forever. Cancel works during this phase too, and the progress
  bar actually moves through it.
- A component refused with **403** means it exists and works but lives in a library this token cannot
  open. That is not the same as a component published nowhere (404), and it is certainly not broken. It
  now gets its own note, with a button to copy the affected component names so access can be requested.
- Retries were made shorter, since Figma answers bursts of lookups without complaint.


## v1.6.3 — 3 September 2026

### Fixed: the token said "active" then "inactive" after a scan
- Two separate faults made the token status untrustworthy:
  - If the check at start-up could not reach the API at all, the plugin assumed the token was fine and
    showed **Token active**. It now shows **Token unverified** in amber and says the scan will settle it.
  - During a scan, a single refused call aborted the whole library verification and declared the token
    dead. One library the token may not read now only affects that library.
- Refusals are no longer guessed at. A component that could not be checked is left out of the results
  instead of being reported as missing, so nothing in the report is a false alarm.
- Figma throttles bursts of REST calls, and a rate-limit reply was being read as an authentication
  failure. Those are now retried with a back-off, and calls are paced to stay under the limit.
- When something genuinely cannot be verified, the banner names the failing endpoint and its HTTP status,
  and the plugin re-checks the token before blaming it, so a permissions problem is never reported as an
  expired token.


## v1.6.2 — 3 September 2026

### Fixed: the health score was clipped
- The score block was placed with absolute positioning and kept the default frame width, so the big
  percentage was cut off on the left. The report header is now a proper auto-layout row, title on the
  left and score on the right, both hugging their content, so nothing can be clipped whatever the numbers.


## v1.6.1 — 3 September 2026

### Health score as the headline
- The canvas report header now leads with the health percentage in 64px bold, right-aligned, replacing the
  decorative bars: a real number makes a better focal point. Under it, "DESIGN SYSTEM HEALTH" and how many
  components are affected out of the total.
- The number is colour-banded: green from 99.5%, amber from 98%, red below. All three verified against the
  dark header background at 7.8:1, 8.2:1 and 5.7:1.
- The panel's health card matches, with the percentage at 44px and the affected-components line beside it.


## v1.6.0 — 3 September 2026

### Severity: "267 to fix" became three honest numbers
- One count treated a dead link and a component that simply wants updating as the same emergency.
  Findings are now graded, and the headline reads for example "5 broken · 12 in components · 250 in designs".
  - **Broken** — the component no longer exists. The only genuinely broken category.
  - **In components** — a live, published component is built on a retired one. Nothing looks wrong, but
    every design placing that component inherits the retired dependency, so one fix here clears them all.
    This is new: the audit now looks at whether an instance sits inside a component definition, and names
    the component it is baked into.
  - **In designs** — a retired component used directly in a design or template. Swap it when convenient.
- The summary shows five cards, the filter chips follow the same three levels, and rows are sorted by
  severity so the urgent work is always at the top.
- "Issues by page" is now **"Issues by live pages"** and states that Graveyard and other retired areas are
  excluded, so it is clear these are the pages the team works on today.
- CSV gains `severity` and `builtInto` columns, and the issue wording explains the difference.
- The canvas report follows the same grading, with a legend that explains all five statuses.


## v1.5.0 — 3 September 2026

### New: "Create report in Figma" replaces the JSON export
- The audit can now build a **formatted report straight onto the canvas**, next to your current content,
  1400px wide like the rest of the generated documentation: dark header with the totals, four summary
  cards, a health bar, issues by page, the full issues table (component, page, location, why it was
  flagged, suggested replacement) and a section explaining what each status means, so the report stands
  on its own when shared.
- Status colours were contrast-checked: amber with near-black text reaches 8.1:1 (AAA); red, green and
  purple with white text reach 5.0:1, 5.4:1 and 5.7:1 (AA at 16px).
- **Copy CSV stays** for spreadsheets and tickets. Download JSON is gone: it was raw data nobody read.
- This button is the only part of the audit that creates anything in your file, and it says so on hover.

### Fixed: tooltips were being cut off
- Status explanations were drawn inside their own chip, so they were clipped at the edge of the panel.
  They now render above everything and shift or flip to stay inside the window.


## v1.4.5 — 3 September 2026

### Fixed: audit results came back blank
- A bug introduced in v1.4.3 left the status tooltip texts inside the wrong function, so drawing the
  results threw an error and the panel stayed empty even though the scan had finished and the tab badge
  showed the issue count. The texts now live at module scope and both views use them.
- The results view is now wrapped in a guard: if drawing ever fails again, the panel explains what
  happened and still offers the raw data, instead of showing nothing.


## v1.4.4 — 3 September 2026

### Loading state for the output panel
- While the plugin reads the file, the output panel now shows a shimmering **code skeleton** instead of an
  empty box, with a rotating status line ("Reading variable collections and their modes", "Pulling the
  global library catalogue", and so on) and a reminder that nothing in the file is being changed.
- It appears on the first automatic sync and on every manual Sync, and disappears as soon as the preview is ready.


## v1.4.3 — 3 September 2026

### Clearer audit results
- Every status now explains itself on hover, plus a one-line legend under the health bar:
  **Linked** (points to a live component), **Deprecated** (properly connected and working, but the
  component has been retired), **Missing** (the component no longer exists, so the link is silently
  broken), **Info only** (cannot be fixed from this page).
- Instances sitting on a **Graveyard page** are no longer counted as issues. Retired material using
  other retired material is expected, and it was more than half of the reported problems.
- The former "Inherited" bucket is now **Info only**, covering both nested instances and Graveyard-page
  ones, and it explains why a single library fix can clear thousands of rows at once.
- Variants of the same component set are grouped into one row, with a variant count, instead of one row
  per variant. Rows are sorted by how many instances are affected.
- CSV and JSON exports follow the same rules, so they match what is on screen.

### Bigger window
- The plugin window opens at 900 × 860 instead of 640 × 760, so the audit table stops truncating names.


## v1.4.2 — 3 September 2026

### Fixed: the audit is now strictly read-only
- The audit no longer imports library components to check whether they are still published.
  `importComponentByKeyAsync` registered every checked component as an "imported component" in the file,
  which counted as a document change and showed up in version history. Library components are now
  verified exclusively through the Figma REST API (published list and Graveyard page per library),
  which is read-only and faster. If the API is unavailable, the report says so instead of guessing.
- Progress steps are now: Scan pages → Verify libraries via API → Build report.


## v1.4.1 — 3 September 2026

### Performance: audit scan is an order of magnitude faster on large files
- One synchronous main-component lookup per instance instead of an awaited call, and no parent walks
  during the scan. Layer paths and nesting are resolved afterwards, only for flagged components
  (about 1% of the instances in the Decibel file). The 187,000-instance library that took over
  15 minutes for 55 pages should now finish in a couple of minutes.
- OK components travel to the panel as a count rather than thousands of instance rows.

### Polish
- Only one sync status indicator: the badge in the header says Synced / Syncing; the Sync button keeps its label and just spins while busy.

### New: cancel, size warning and last-scan memory
- **Cancel** button while scanning.
- Before scanning, a **size warning** shows the number of pages and how long the last full scan took,
  suggesting "Current page" for a quick check.
- The Audit tab remembers the **last scan of each file** (per user, newest three files): when you
  reopen the plugin it shows when it ran, how long it took, the health percentage and the four counts,
  with **View results** to reopen the previous report (Locate still works for layers that still exist)
  and **Scan again**.


## v1.4.0 — 3 September 2026

### New: Audit tab — broken links and deprecated components
- A third main tab, **Audit**, scans every instance in the file (current page or all pages) and flags:
  - **Deprecated** — the main component lives in a library page whose name contains "Graveyard", or its
    description carries an `@deprecated` tag (optionally with a replacement, e.g. `@deprecated → use Button/Primary`),
    or its name contains "DEPRECATED".
  - **Missing** — the main component was deleted, or its key is no longer published in any accessible library
    (a silent broken link: Figma still renders the instance from its local cache).
  - **Inherited** — instances nested inside another instance, listed in a separate collapsed section and not
    counted as issues. Fixing them belongs to the component that contains them.
- Results are visual: a health bar with four tiles (Linked / Deprecated / Missing / Inherited), issues per page,
  filter chips, and an issues table grouped by component with **Locate** (selects the layer on canvas and
  switches page if needed). Export as CSV or JSON with component, page, layer path, issue type, status,
  detection source, expected replacement and a deep link.
- Library components whose key cannot be imported are settled through the Figma REST API: the plugin discovers the
  libraries in use, reads each library's Graveyard page and tells "unpublished in Graveyard" apart from "deleted".
  Works across several libraries (e.g. Decibel Mobile App and Decibel Global Branding).
- The audit is read-only: nothing in the file is ever modified.

### API token
- The plugin ships with a shared Figma access token injected at packaging time (`build-zip.sh`); the repository
  only ever contains a placeholder. The Audit header shows the token status and days left before expiry
  (green, amber at 14 days, red at 3). A settings card allows an override token, stored on that computer only.
- Without a valid token the scan still runs, but unpublished Graveyard components are reported as Missing.

### Manifest
- Plugin name updated to **V.1.4.0**; `https://api.figma.com` added to the allowed network domains.


## v1.3.3 — 21 August 2026

### Fixed: font-weight bound as a number variable now shows its token and raw value
- Figma lets a text style's weight be bound under two different API keys: `fontStyle` (a string variable, e.g. "Semi Bold") or `fontWeight` (a number variable, e.g. 600).
- The Responsive values table and the property badges only checked `fontStyle`, so styles bound via `fontWeight` showed the literal name (e.g. "SemiBold") with no token or raw value.
- Both keys are now checked — the font-weight row shows the bound token (e.g. `global-primitive/type/font-weight/body/semi-bold`) with its numeric value (600) in the chip below.

## v1.3.2 — 21 August 2026

### Improved: raw values as chips, and single-mode bindings resolved
- Raw values in the Responsive values table now render as a **discreet chip** (pink text #C2185B on a light pink tint #FCE7EF — contrast 5.0:1, WCAG AA at 16px) instead of plain pink text.
- Properties bound **directly to single-mode collections** (for example primitives such as `global-primitive/type/font-family/body`) now resolve too: the chip shows the concrete value (for example `AJ Bell Platform Text`), repeated across breakpoints since it does not vary.
- A chip is skipped when it would duplicate the text above it (literal values are already raw).

### Improved: description callout and breathing room
- Style descriptions now render as a **highlighted callout** — light blue block (#E8F1FB) with a dark left accent bar — instead of plain grey text, so they no longer get lost between the badges. Contrast 12.8:1, WCAG AAA at 16px. Applied to text, colour and effect rows.
- More vertical spacing between the elements of each row (name, description, badges, responsive table), so the content no longer feels cramped.

## v1.3.1 — 21 August 2026

### Improved: raw values in the Responsive values table
- Each cell that shows a token alias now also shows the **resolved raw value** underneath, in a pink label (for example `global-primitive/type/font-size/3xl` with `40` below it).
- Works for numbers, strings, booleans and colours (colours show their hex code).
- Pink label colour is #C2185B — contrast 5.9:1 on the white table background, passing WCAG AA at 16px.
- Cells with literal (unbound) values are unchanged, as they already show the raw value.

## v1.3.0 — 21 August 2026

### New: Responsive values table in Text Style documentation
- Each Text Style row in the generated documentation now includes a **"Responsive values" table** when the style is bound to a multi-mode variable collection (for example the responsive collection with Mobile Small, Mobile Large, Tablet and Desktop).
- The table shows **all standard text properties**: font-size, line-height, font-family, font-weight, letter-spacing and paragraph-spacing.
- Cells show the **alias name** per breakpoint (for example `global-primitive/type/font-size/m`), matching what designers see in Figma's variable collection view. Properties without a multi-mode binding show their literal value repeated across all columns.
- Numbers are rounded to a maximum of 3 decimals.

### New: Global (Team Library) variable export
- The plugin now pulls the **entire global library catalogue** — every collection and variable published by the libraries enabled in the file, used or not — via the Team Library API.
- Global collections are marked with a purple **GLOBAL badge** in the plugin panel.
- Exports label global content clearly:
  - CSS: comment shows `(global — LibraryName)` next to the collection name.
  - Flutter: global collections get a `Global` suffix in the class name.
- Variables are imported in **parallel batches of 25**, bringing a full library of ~537 variables from ~90 seconds down to ~8 seconds.
- Alias chains that cross **remote/library collections** are now resolved (async, cached, up to 10 levels deep).
- This layer is fully additive and wrapped in try/catch — a failure in the global discovery can never break the existing local export.

### Manifest
- Plugin name updated to **V.1.3.0**.
- Added the **`teamlibrary` permission**, required by the Team Library API.

### Housekeeping
- `dist/` re-synchronised as an exact mirror of the root files.
- Added developer test scripts (`test-remote-vars.js`, `test-formatters.js`) — not part of the packaged plugin.

---

## v1.2.0 — 25 February 2026
- Publish to Bitbucket: 4-step wizard (settings, review, publishing, done), branch creation/selection, PR title/description, reviewer picker, GitHub-style code compare.
- Alias chain resolution up to 10 levels deep (`resolveAliasChain`).
- Text styles export their bound variables; documentation badges show real variable names.
- Removed comment headers from the Flutter export output.

## v1.0 — 22 February 2026
- Initial release: extract styles and variables, on-canvas documentation generator, exports in JSON, CSS, Flutter and W3C DTCG, Live Sync.
