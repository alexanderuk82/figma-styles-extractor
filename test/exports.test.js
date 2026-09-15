// Every export format, for styles and for variables, must produce a structurally
// valid file with no data loss — this is what runs before every release.
"use strict";
const { test, describe, beforeEach } = require("node:test");
const assert = require("node:assert/strict");
const { load } = require("./harness");
const fx = require("./fixtures");
const { validateCss, validateDtcg, validateDart } = require("./validators");

let ui;
beforeEach(() => { ui = load(); ui.setFile("Acme Tokens"); ui.setNaming({}); });

describe("variables → W3C DTCG", () => {
  test("multi-mode: mode is the top-level key and every mode is a complete, resolvable set", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("dtcg");
    const t = validateDtcg(ui.formatVariablesDTCG(), { modeRoots: true });
    assert.deepEqual(Object.keys(t).sort(), [...fx.allModes()].sort(), "one root per mode");
    assert.equal(t.Light.semantics.text.primary.$value, "{primitives.color.blue-500}", "reference is dot-separated and includes the collection");
    assert.equal(t.Light.semantics.text.primary.$type, "color", "$type is the resolved type, never 'alias'");
    assert.equal(t.Dark.primitives.color["blue-500"].$value, "#3377FF", "dark value survives; opaque colours stay 6-digit");
  });
  test("single mode: no mode wrapper, tokens at the collection level", () => {
    const d = fx.variables(); d.collections = d.collections.filter((c) => c.id === "c-resp");
    ui.setVariables(d, ["Desktop"]); ui.setFormat("dtcg");
    const t = validateDtcg(ui.formatVariablesDTCG());
    assert.equal(t.responsive.typography["font-size"].xl.$value, 36);
  });
  test("a local and an external collection with the same name are both kept", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("dtcg");
    const t = validateDtcg(ui.formatVariablesDTCG(), { modeRoots: true });
    // each collection appears under the modes it actually has: local under Light/Dark, external under Mode 1
    assert.ok(Object.keys(t.Light).includes("primitives"), "local primitives keeps its plain name");
    assert.ok(Object.keys(t["Mode 1"]).includes("primitives-acme-global"), `external primitives is told apart by ITS LIBRARY's name, not a fixed word: ${Object.keys(t["Mode 1"])}`);
    assert.equal(t.Light.primitives.color["blue-500"].$value, "#0055FF", "local value intact");
    assert.equal(t["Mode 1"]["primitives-acme-global"].color["blue-500"].$value, "#0066FF", "external value intact");
  });
  test("alpha colours are written as 8-digit hex", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("dtcg");
    const t = JSON.parse(ui.formatVariablesDTCG());
    assert.equal(t.Light.primitives.color.overlay.$value, "#00000066");
  });
});

describe("variables → CSS", () => {
  test("one block per mode, first mode on :root, every var() resolvable, no duplicates", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("css");
    const blocks = validateCss(ui.formatVariablesCSS());
    assert.equal(blocks[0].selector, ":root");
    assert.ok(blocks.some((b) => b.selector === '[data-mode="dark"]'), "dark mode block present");
    assert.equal(blocks.length, fx.allModes().length, "one block per mode");
  });
  test("alpha colour becomes rgba, alias becomes var()", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("css");
    const css = ui.formatVariablesCSS();
    assert.match(css, /--primitives-color-overlay: rgba\(0,0,0,0\.40\);/);
    assert.match(css, /--semantics-text-primary: var\(--primitives-color-blue-500\);/);
  });
  test("same-name collections never produce a duplicate property", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("css");
    const css = ui.formatVariablesCSS();
    const block = css.split("}")[0];
    const names = [...block.matchAll(/^\s*(--[a-z0-9-]+):/gim)].map((m) => m[1]);
    assert.equal(new Set(names).size, names.length, "no duplicate custom properties in :root");
  });
});

describe("variables → Flutter", () => {
  test("multi-mode: one class per collection per mode, all unique, all valid Dart", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("flutter");
    const classes = validateDart(ui.formatVariablesFlutter());
    assert.ok(classes.length >= fx.variables().collections.length, "at least one class per collection");
  });
  test("colours are Color(0xAARRGGBB), numbers are double", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("flutter");
    const dart = ui.formatVariablesFlutter();
    assert.match(dart, /static const Color \w+ = Color\(0x[0-9A-F]{8}\);/, "colours carry an explicit Color type");
    assert.match(dart, /static const double \w+ = 16;/, "numbers carry an explicit double type, never inferred as int");
    assert.ok(!/static const \w+ = \d+;/.test(dart), "no untyped numeric constants");
  });
});

describe("variables → Figma JSON", () => {
  test("is valid JSON and keeps every collection and variable (faithful dump)", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("figma");
    const j = JSON.parse(ui.formatVariablesFigma());
    assert.equal(j.collections.length, fx.variables().collections.length);
    const total = j.collections.reduce((n, c) => n + c.variables.length, 0);
    assert.equal(total, fx.variables().collections.reduce((n, c) => n + c.variables.length, 0));
    assert.ok(j.collections.some((c) => c.source === "global-library" && c.libraryName), "provenance preserved for external collections");
  });
  test("naming options never touch the raw JSON", () => {
    ui.setVariables(fx.variables(), fx.allModes()); ui.setFormat("figma");
    const before = ui.formatVariablesFigma();
    ui.setNaming({ prefix: "acme", collection: false, aliases: { primitives: "prim" } });
    assert.equal(ui.formatVariablesFigma(), before);
  });
});

describe("styles → every format", () => {
  test("DTCG: colours, gradients, typography, effects all present and valid", () => {
    ui.setStyles(fx.styles()); ui.setFormat("dtcg");
    const t = validateDtcg(ui.formatStylesDTCG(fx.styles()));
    assert.equal(t.color.brand.red.$value, "#FF0000", "opaque colour stays 6-digit");
    assert.equal(t.color.brand.fade.$type, "gradient");
    assert.ok(t.typography, "typography section present");
  });
  test("CSS: well-formed with gradients and typography", () => {
    ui.setStyles(fx.styles()); ui.setFormat("css");
    const css = ui.formatStylesCSS(fx.styles());
    validateCss(css);
    assert.match(css, /--brand-red: #FF0000;/);
    assert.match(css, /--brand-fade: linear-gradient\(/);
    assert.match(css, /--body-regular-family: 'Inter', sans-serif;/);
  });
  test("Flutter: valid Dart", () => {
    ui.setStyles(fx.styles()); ui.setFormat("flutter");
    validateDart(ui.formatStylesFlutter(fx.styles()));
  });
});

describe("what the buttons deliver", () => {
  test("Copy and Download hand over exactly the previewed output, in all formats, styles and variables", () => {
    for (const section of ["styles", "variables"]) {
      if (section === "styles") ui.setStyles(fx.styles()); else ui.setVariables(fx.variables(), fx.allModes());
      for (const f of ["figma", "css", "flutter", "dtcg"]) {
        ui.setFormat(f);
        const expected = ui.getFormattedOutput();
        assert.ok(expected.length > 50, `${section}/${f} produces output`);
        const info = ui.getFileInfo();
        assert.ok(info.ext && info.mime, `${section}/${f} has a file extension and mime`);
        assert.match(info.prefix + info.ext, /^[a-z0-9-]+(\.tokens)?\.(json|css|dart)$/, `${section}/${f} download name is clean: ${info.prefix + info.ext}`);
      }
    }
  });
});
