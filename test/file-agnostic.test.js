// The plugin runs in ANY file: every default derives from the open file, nothing
// is hardcoded to a project, and the code carries no brand words.
"use strict";
const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { load } = require("./harness");
const fx = require("./fixtures");

describe("defaults derive from the open file", () => {
  test("namespace suggestion strips only structural noise", () => {
    const ui = load();
    assert.equal(ui.auFileSlug("[Brand] Acme Global Branding"), "acme-global-branding");
    assert.equal(ui.auFileSlug("[Design System] Acme Mobile App (ver2.0)"), "acme-mobile-app");
    assert.equal(ui.auFileSlug("Web Kit v3"), "web-kit");
    assert.equal(ui.auFileSlug("Iconos Banco Azul"), "iconos-banco-azul");
    assert.equal(ui.auFileSlug(""), "tokens");
  });
  test("publish path, branch, download name and Flutter class all carry the file, so two files never collide", () => {
    const ui = load(); ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("dtcg");
    const seen = new Set();
    for (const f of ["[Brand] Acme Global Branding", "[Design System] Acme Mobile App (ver2.0)", "Checkout Flow"]) {
      ui.setFile(f);
      const p = ui.pubDefaultFilePath("variables", "dtcg"), b = ui.pubDefaultBranch("variables"), d = ui.getFileInfo().prefix, c = ui.getFlutterClassName();
      for (const v of [p, b, d, c]) { assert.ok(!seen.has(v), `collision across files: ${v}`); seen.add(v); }
      assert.match(p, /^ds\/tokens\/[a-z0-9-]+-variables\.tokens\.json$/);
      assert.match(c, /^[A-Z][A-Za-z0-9]*Tokens$/); assert.doesNotMatch(c, /TokensTokens/, "no doubled suffix");
    }
  });
  test("a file whose name already ends in Tokens does not get TokensTokens", () => {
    const ui = load(); ui.setFile("Acme Tokens"); ui.setVariables(fx.variables(), fx.allModes());
    assert.equal(ui.getFlutterClassName(), "AcmeTokens");
  });
});

describe("no project knowledge in the source", () => {
  test("code and comments carry no brand or file names", () => {
    const root = path.join(__dirname, "..");
    const banned = /decibel|aj ?bell|global branding|mobile app|testing_/i;
    for (const f of ["code.js", "ui.html"]) {
      const src = fs.readFileSync(path.join(root, f), "utf8").replace(/ds-styles-extractor-ajbell/g, "");
      const hits = src.split("\n").map((l, i) => (banned.test(l) ? `${f}:${i + 1}: ${l.trim().slice(0, 80)}` : null)).filter(Boolean);
      assert.deepEqual(hits, [], "project-specific words found");
    }
  });
  test("the only name-based lookups are the plugin's own artefacts or non-binding preferences", () => {
    const root = path.join(__dirname, "..");
    const src = fs.readFileSync(path.join(root, "code.js"), "utf8") + fs.readFileSync(path.join(root, "ui.html"), "utf8");
    const cmp = [...src.matchAll(/\.name(?:\.toLowerCase\(\))?\s*===?\s*['"]([a-zA-Z][^'"]*)['"]/g)].map((m) => m[1]);
    const allowed = new Set(["Documentation", "components", "main", "master"]);
    assert.deepEqual(cmp.filter((n) => !allowed.has(n)), [], "unexpected hardcoded name comparison");
  });
});

describe("start-up never hides local data", () => {
  test("first delivery with libraries pending keeps the output visible and counts populated", () => {
    const ui = load();
    ui.onmessage({ type: "all-data", payload: { styles: fx.styles(), variables: fx.variables() }, globalsPending: true });
    assert.equal(ui.store["output"].classList.contains("loading"), false, "output not hidden behind the skeleton");
    assert.equal(ui.store["badge-styles"].textContent, 4);
    assert.match(ui.store["sync-label"].textContent, /loading external libraries/);
    ui.onmessage({ type: "globals-done", ok: true, reason: "none" });
    assert.equal(ui.store["sync-label"].textContent, "Synced");
  });
});

describe("markup hygiene", () => {
  test("no JavaScript unicode escapes are left inside HTML attributes (they render literally)", () => {
    const html = require("fs").readFileSync(require("path").join(__dirname, "..", "ui.html"), "utf8");
    const markup = html.slice(0, html.indexOf("<script>"));   // attributes live in the markup, not the script
    const hits = [...markup.matchAll(/(data-tip|title|placeholder|alt)="[^"]*\\u[0-9a-fA-F]{4}[^"]*"/g)].map((m) => m[0].slice(0, 70));
    assert.deepEqual(hits, [], "literal \\uXXXX inside an HTML attribute");
  });
});
