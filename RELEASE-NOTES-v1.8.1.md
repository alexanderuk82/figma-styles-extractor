# DS Styles Extractor v1.8.1 — what's new

Everything since the version you had, v1.3.3. The full engineering history is in `CHANGELOG.md`.

## New: Audit tab

A third tab that scans every instance in the file and finds components that are broken or retired.
A scan is **read-only**: nothing in your file is created, changed or published by it.

Findings are graded, so you always know what actually needs attention:

- **Broken** — the component no longer exists, or is no longer published in any library you can reach.
  Figma still draws it from its local cache, so it looks fine, but the link is dead and it will never
  update again. This is the only category that is genuinely broken.
- **In components** — a live, published component is built on a retired one. Nothing looks wrong, but
  every design placing that component inherits the retired dependency, so one fix here clears them all.
  The report names the component it is baked into.
- **In designs** — a retired component used directly in a design or template. Swap it when convenient.
- **Info only** — nested inside another instance, or sitting on a Graveyard page. Not counted as issues.

A component counts as retired when it sits in a page whose name contains "Graveyard", carries an
`@deprecated` tag in its description, or says DEPRECATED in its name. An `@deprecated → use Button/Primary`
tag also fills in the suggested replacement.

Private components, the ones whose name starts with a dot or an underscore, are skipped and reported as
such. They are the internal parts of published components, they are never published themselves, and they
do not affect the health score.

### Using it

- **Scan** the current page or the whole file. Progress is shown throughout and can be cancelled.
- **Locate** jumps to any layer, switching page if needed.
- **Copy CSV** for tickets and spreadsheets, with severity, component, page, layer path, the reason it was
  flagged and a deep link to each layer.
- **Create report in Figma** builds a formatted report on the canvas: health score, summary, issues by live
  page, the full table and a legend explaining every status. This is the only button that creates anything
  in your file, and it says so.
- The plugin remembers the last scan of each file, so reopening it shows the previous result straight away.
- Any note that appears explains what it means, why it happened and whether it affects the score, and lets
  you list the components it is talking about.

## New: you choose how exported token names are built

A **Naming** bar appears above the format tabs whenever CSS or W3C DTCG is selected on the Variables tab:

- **Collection prefix** on or off.
- **Merge repeated words**, so `global-global` becomes `global`.
- **Prefix** of your own, for example `ds`.
- **Per collection** gives each collection a different name in the export: `primitives` can go out as
  `prim`, `secondary` as `sec`. Aliases follow the renames, so references never break.
- A live example shows exactly what the next token name will look like, and the settings are remembered
  after closing the plugin and restarting Figma.

The raw Figma JSON is deliberately left untouched: it is a faithful dump of the file, used for diffs when
publishing. W3C DTCG is the format to pick for clean names.

## Fixed

- **The plugin took minutes to open on a large design system.** The panel no longer waits for the global
  library: local styles and variables appear at once and the library arrives as a second update. The
  library itself now loads in parallel rather than one variable at a time.
- **CSS variables from two collections with the same name overwrote each other.** A local *primitives* and
  a global *primitives* both produced `--primitives-…` names and the global silently won.
- **The W3C DTCG export was not usable by a token pipeline.** References were written the Figma way and
  without the collection, so they never resolved; `$type` said `alias`, which is not a type; and modes were
  nested inside the token path, making "Light" part of every token's name. The mode is now the top level,
  so splitting the file into light.json and dark.json is just taking one key.
- **Pull requests were pre-filled with a made-up ticket number.**

## Note for the team

The audit needs a Figma API token to tell "unpublished in a Graveyard" apart from "deleted". A shared token
is built into this release, so it works out of the box — the header shows **Token active**. If it ever shows
otherwise, tell Alex; he rotates it and publishes a new build. You can enter your own token from that same
control if you prefer.
