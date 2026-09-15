// The Naming bar: namespace, collection prefix, merged words, per-collection
// aliases — and the promise that references always follow the renames.
"use strict";
const { test, describe, beforeEach } = require("node:test");
const assert = require("node:assert/strict");
const { load } = require("./harness");
const fx = require("./fixtures");
const { validateCss, validateDtcg, validateDart } = require("./validators");

let ui;
beforeEach(() => { ui = load(); ui.setFile("Acme Tokens"); ui.setVariables(fx.variables(), fx.allModes()); });

describe("namespace", () => {
  test("becomes the DTCG top-level key, the CSS leading segment and the Dart class prefix", () => {
    ui.setNaming({ prefix: "acme" });
    ui.setFormat("dtcg"); const t = validateDtcg(ui.formatVariablesDTCG(), { modeRoots: true });
    assert.deepEqual(Object.keys(t.Light), ["acme"], "DTCG root is the namespace");
    ui.setFormat("css"); assert.match(ui.formatVariablesCSS(), /^\s*--acme-primitives-color-blue-500:/m);
    ui.setFormat("flutter"); assert.ok(validateDart(ui.formatVariablesFlutter()).every((c) => c.startsWith("Acme")), "every Dart class starts with the namespace");
  });
  test("two files with identical tokens do not collide once merged", () => {
    ui.setFormat("dtcg");
    ui.setNaming({ prefix: "brand" }); const a = JSON.parse(ui.formatVariablesDTCG());
    ui.setNaming({ prefix: "app" }); const b = JSON.parse(ui.formatVariablesDTCG());
    const merged = { ...a.Light, ...b.Light };
    assert.deepEqual(Object.keys(merged).sort(), ["app", "brand"]);
    assert.equal(merged.brand.primitives.color["blue-500"].$value, merged.app.primitives.color["blue-500"].$value, "both copies survive the merge");
  });
  test("a multi-word namespace is split into path segments", () => {
    ui.setNaming({ prefix: "acme design" }); ui.setFormat("dtcg");
    const t = JSON.parse(ui.formatVariablesDTCG());
    assert.ok(t.Light.acme && t.Light.acme.design, "acme → design nesting");
  });
});

describe("collection prefix and merged words", () => {
  test("collection prefix off drops the collection from every name, and references still resolve", () => {
    ui.setNaming({ collection: false });
    ui.setFormat("css"); const css = ui.formatVariablesCSS(); validateCss(css);
    assert.match(css, /^\s*--color-blue-500:/m); assert.match(css, /--text-primary: var\(--color-blue-500\);/);
    ui.setFormat("dtcg"); const t = validateDtcg(ui.formatVariablesDTCG(), { modeRoots: true });
    assert.equal(t.Light.text.primary.$value, "{color.blue-500}");
  });
  test("merged words collapse an immediate repeat, and only that", () => {
    const d = fx.variables(); d.collections = [{ id: "c-brand", name: "brand", modes: [{ id: "M", name: "M" }], variables: [
      { id: "1", name: "brand-primitive/type/font-size/xl", resolvedType: "FLOAT", valuesByMode: { M: { type: "number", value: 32 } } } ] }];
    ui.setVariables(d, ["M"]); ui.setFormat("css");
    ui.setNaming({ dedupe: false }); assert.match(ui.formatVariablesCSS(), /--brand-brand-primitive-type-font-size-xl/);
    ui.setNaming({ dedupe: true }); assert.match(ui.formatVariablesCSS(), /--brand-primitive-type-font-size-xl/);
    assert.doesNotMatch(ui.formatVariablesCSS(), /--brand-primitive-type-size-xl/, "non-adjacent repeats are left alone");
  });
});

describe("per-collection aliases", () => {
  test("rename the collection in every format and references follow", () => {
    ui.setNaming({ aliases: { primitives: "prim", semantics: "sem" } });
    ui.setFormat("css"); const css = ui.formatVariablesCSS(); validateCss(css);
    assert.match(css, /--sem-text-primary: var\(--prim-color-blue-500\);/);
    ui.setFormat("dtcg"); const t = validateDtcg(ui.formatVariablesDTCG(), { modeRoots: true });
    assert.equal(t.Light.sem.text.primary.$value, "{prim.color.blue-500}");
    ui.setFormat("flutter"); const classes = validateDart(ui.formatVariablesFlutter());
    assert.ok(classes.some((c) => /Prim/.test(c)) && classes.some((c) => /Sem/.test(c)), `Dart classes use the aliases: ${classes}`);
  });
  test("a blank alias means the real name", () => {
    ui.setNaming({ aliases: { primitives: "   " } }); ui.setFormat("dtcg");
    assert.ok(JSON.parse(ui.formatVariablesDTCG()).Light.primitives, "blank alias ignored");
  });
});

describe("name clashes", () => {
  test("only a genuinely duplicated collection name is suffixed, and with its library's name", () => {
    ui.setNaming({}); ui.setFormat("dtcg");
    const t = JSON.parse(ui.formatVariablesDTCG());
    assert.ok(t.Light.primitives && t.Light.semantics && t.Desktop.responsive, "local names untouched (each under its own modes)");
    assert.ok(t["Mode 1"]["primitives-acme-global"], "external duplicate carries the library slug");
    assert.equal(t.Light["primitives-acme-global"], undefined, "suffix never leaks onto the local one");
  });
  test("Dart: several collections always keep the collection in the class name, even with the prefix off", () => {
    ui.setNaming({ prefix: "acme", collection: false }); ui.setFormat("flutter");
    const classes = validateDart(ui.formatVariablesFlutter());   // validateDart asserts uniqueness
    assert.ok(classes.length > 1);
  });
});

describe("styles follow the namespace too", () => {
  test("DTCG root and CSS prefix for styles; nothing changes without a namespace", () => {
    ui.setStyles(fx.styles());
    ui.setNaming({}); const plain = ui.formatStylesDTCG(fx.styles());
    assert.ok(JSON.parse(plain).color, "no wrapper without a namespace");
    ui.setNaming({ prefix: "acme" });
    const t = JSON.parse(ui.formatStylesDTCG(fx.styles())); assert.ok(t.acme.color.brand.red, "namespace wraps the styles");
    const css = ui.formatStylesCSS(fx.styles()); validateCss(css); assert.match(css, /--acme-brand-red: #FF0000;/);
    assert.doesNotMatch(css, /--acme-acme-/, "prefix applied exactly once");
  });
});

describe("Flutter constants honour the naming options", () => {
  test("with 'merge repeated words' OFF the constant names are exactly the historical toCamel output", () => {
    ui.setNaming({ dedupe: false }); ui.setFormat("flutter");
    const dart = ui.formatVariablesFlutter();
    assert.match(dart, /static const Color colorBlue500 =/);
    assert.match(dart, /static const double space4 =/);
  });
  test("with it ON an immediately repeated word is kept once, and nothing else moves", () => {
    const d = fx.variables(); d.collections = [{ id: "c", name: "brand", modes: [{ id: "M", name: "M" }], variables: [
      { id: "1", name: "color/color-red", resolvedType: "COLOR", valuesByMode: { M: { type: "color", hex: "#F00", r: 255, g: 0, b: 0, a: 1 } } },
      { id: "2", name: "space/4", resolvedType: "FLOAT", valuesByMode: { M: { type: "number", value: 4 } } } ] }];
    ui.setVariables(d, ["M"]); ui.setFormat("flutter");
    ui.setNaming({ dedupe: false }); assert.match(ui.formatVariablesFlutter(), /colorColorred =/);
    ui.setNaming({ dedupe: true });  const on = ui.formatVariablesFlutter();
    assert.match(on, /colorRed =/, "color/color-red → colorRed"); assert.match(on, /space4 =/, "untouched name stays");
    validateDart(on);
  });
  test("the hint tells the user when a switch changes nothing for this format", () => {
    ui.setFormat("flutter"); ui.setNaming({});
    ui.get("auNamingRefresh(false)");
    const hint = ui.store["nb-hint"].textContent;
    assert.match(hint, /Collection prefix: no effect here/, "8 collections → the collection stays in Dart class names");
    assert.match(hint, /Dart keeps the collection/);
    ui.setFormat("css"); ui.get("auNamingRefresh(false)");
    assert.doesNotMatch(ui.store["nb-hint"].textContent, /Collection prefix: no effect/, "in CSS the prefix does change the output");
  });
});
