# DS Styles Extractor v1.10.0 — what's new

If you are on **v1.8.1**, read the first section. If you are still on **v1.3.3**, everything below applies.
The full engineering history is in `CHANGELOG.md`.

## New since v1.8.1

### The plugin works in any file
- Opened in the brand library, the plugin used to stay on "Reading your file…" and its variables
  documentation refused to run because it expected a collection called "Components". Both are fixed.
  Any collection can be documented, and local styles and variables appear at once while external
  libraries load in the background.
- Nothing in the plugin assumes which file it runs in. Every default — namespace, publish path, branch
  name, download name, Flutter class name — is derived from the file that is open, and can be edited.
- Collections that come from a library this file uses are badged **EXTERNAL**, with the source library
  named beside them. Where two collections share a name, the external one is told apart by its library's
  name.

### Exports from two files no longer overwrite each other
- **Naming settings are stored per file.** The brand library keeps its own namespace and aliases; the app
  keeps its own.
- The **Namespace** goes in front of every token from a file: the first word of each CSS custom property,
  the top-level key in W3C DTCG, the start of every Flutter class name. Two files with different
  namespaces can be merged by a build with nothing overwritten. "Suggest from file" proposes one from the
  file title; edit it to anything you like.
- The default publish path carries the file name too, so two files never default to the same destination.
- The namespace covers styles as well as variables.

### Flutter
- **Constant style**, remembered per file: **camelCase** (the Dart convention, now the default),
  **snake_case**, or **As before** — the names earlier exports produced, kept so existing code compiles
  until you choose to rename. The preview shows the new names before any export, and the publish
  wizard's Code Compare lists every renamed constant before a pull request is created.
- Every constant carries an explicit type (`Color`, `double`, `String`, `bool`). Without it, a number was
  inferred as `int` and could not be passed where Flutter expects a `double`.
- The Naming switches now reach Flutter, and a line under the bar explains when a switch cannot change the
  current export (for example, with several collections selected Dart keeps the collection in each class
  name so they stay unique).

### Quality
- A test suite runs against the real plugin code before every release: every export format for styles
  and variables, multi-mode output, alias resolution, naming options, name clashes, the audit model, and
  the rule that nothing is hardcoded to a project. No release is built while any test fails.
- Every Naming control has a tooltip saying what it does in each format.

## Since v1.3.3 — the Audit tab

A third tab that scans every instance in the file and finds components that are broken or retired.
A scan is **read-only**: nothing in your file is created, changed or published by it.

Findings are graded:
- **Broken** — the component no longer exists, or is no longer published in any library you can reach.
  Figma still draws it from its local cache, so it looks fine, but the link is dead. The only category
  that is genuinely broken.
- **In components** — a live, published component is built on a retired one. Every design placing that
  component inherits the retired dependency, so one fix here clears them all.
- **In designs** — a retired component used directly in a design or template. Swap it when convenient.
- **Info only** — nested inside another instance, or on a Graveyard page. Not counted as issues.

A component counts as retired when it sits in a page whose name contains "Graveyard", carries an
`@deprecated` tag in its description, or says DEPRECATED in its name. Private components (names starting
with a dot or underscore) are skipped and do not affect the health score.

Scan the current page or the whole file, cancel at any time, **Locate** any layer, **Copy CSV**, or
**Create report in Figma** — the one action that creates anything in your file. The plugin remembers the
last scan of each file.

## Also since v1.3.3
- Text style documentation shows a **Responsive values** table with the bound token per breakpoint and its
  resolved value; style descriptions render as a highlighted callout.
- Larger window; loading skeleton while the file is read; the plugin opens immediately instead of waiting
  for external libraries.
- W3C DTCG is now usable by a token pipeline: dot-separated references that include the collection, real
  `$type` values, and modes at the top level so light.json / dark.json is one key each.
- Pull requests no longer start with a made-up ticket number.

## Note for the team
The audit needs a Figma API token to tell "unpublished in a Graveyard" apart from "deleted". A shared
token is built into this release, so it works out of the box — the header shows **Token active**. If it
ever shows otherwise, tell Alex; he rotates it and publishes a new build. You can enter your own token
from that same control if you prefer.
