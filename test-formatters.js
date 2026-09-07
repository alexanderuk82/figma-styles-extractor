// End-to-end test of the REAL export formatters in ui.html.
// Loads the actual <script> into a vm with a permissive DOM stub, injects test
// data with a LOCAL + GLOBAL collection sharing the name "primitives" (the real
// collision case), and asserts every format keeps both without data loss.
// Run with: node test-formatters.js
"use strict";
const fs = require("fs");
const vm = require("vm");

const html = fs.readFileSync(require("path").join(__dirname, "ui.html"), "utf8");
const script = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1])[0];

// ── Permissive DOM/browser stubs ──
function makeStub() {
  const classList = { add() {}, remove() {}, toggle() {}, contains() { return false; } };
  return new Proxy(function () {}, {
    get(t, prop) {
      if (prop === "classList") return classList;
      if (prop === "style" || prop === "dataset") return {};
      if (prop === "querySelector" || prop === "closest") return () => makeStub();
      if (prop === "querySelectorAll") return () => [];
      if (prop === "appendChild") return (x) => x;
      if (prop === "children") return [];
      if (["value", "textContent", "innerHTML", "className", "id", "name", "type"].includes(prop)) return "";
      if (typeof prop === "symbol") return undefined;
      return makeStub(); // chainable / callable / settable
    },
    set() { return true; },
    apply() { return makeStub(); },
  });
}
const document = {
  getElementById: () => makeStub(),
  querySelector: () => makeStub(),
  querySelectorAll: () => [],
  createElement: () => makeStub(),
  addEventListener: () => {},
  body: makeStub(),
};
const sandbox = {
  document, parent: { postMessage() {} }, navigator: {}, console,
  setTimeout: () => 0, setInterval: () => 0, clearInterval: () => {}, clearTimeout: () => {},
  Date, JSON, Math, Object, Array, String, Number, Boolean, Set, Map, RegExp, isNaN, parseInt, parseFloat, URL: undefined, Blob: undefined,
};
sandbox.window = sandbox;
sandbox.globalThis = sandbox;
vm.createContext(sandbox);

// Append a runner in the SAME lexical scope so it can set the `let` globals.
const runner = `
;globalThis.__run = function (td) {
  variablesData = td.variablesData;
  selectedCollections = td.selectedCollections;
  selectedModes = td.selectedModes;
  allModes = td.allModes;
  varSearchQuery = td.search || '';
  currentSection = 'variables';
  currentFormat = 'dtcg';
  return {
    dtcg: formatVariablesDTCG(), flutter: formatVariablesFlutter(),
    css: formatVariablesCSS(), figma: formatVariablesFigma(),
    matchCount: varSearchMatchCount(), filteredCols: getFilteredVariables().map(c => c.name)
  };
};
;globalThis.__collapse = function (v) { setCollectionsCollapsed(v); return collectionsCollapsed; };`;

try { vm.runInContext(script + runner, sandbox); }
catch (e) { console.error("✗ failed to load ui.html script:", e.message); process.exit(1); }

// ── Test data: LOCAL "primitives" + GLOBAL "primitives" (name collision) ──
const data = {
  variablesData: { _meta: {}, collections: [
    { id: "L1", name: "primitives", modes: [{ id: "m1", name: "Mode 1" }], variables: [
      { id: "lv1", name: "color/bg", resolvedType: "COLOR", valuesByMode: { "Mode 1": { type: "color", hex: "#ffffff", a: 1, r: 255, g: 255, b: 255 } } },
      { id: "lv2", name: "components/skeleton/base", resolvedType: "COLOR", valuesByMode: { "Mode 1": { type: "color", hex: "#eeeeee", a: 1, r: 238, g: 238, b: 238 } } },
    ] },
    { id: "G1", name: "primitives", source: "global-library", libraryName: "Brand Global", remote: true, modes: [{ id: "m1", name: "Mode 1" }], variables: [
      { id: "gv1", name: "global-primitive/color/brand", resolvedType: "COLOR", valuesByMode: { "Mode 1": { type: "color", hex: "#0066ff", a: 1, r: 0, g: 102, b: 255 } } },
    ] },
  ] },
  selectedCollections: new Set(["L1", "G1"]),
  selectedModes: new Set(),
  allModes: ["Mode 1"],
};
const out = sandbox.__run(data);

let pass = 0, fail = 0;
function ok(n, c, e) { if (c) { pass++; console.log("  ✅ " + n); } else { fail++; console.log("  ❌ " + n + (e ? "  → " + e : "")); } }

console.log("\n── DTCG (was: global overwrote local) ──");
const dtcg = JSON.parse(out.dtcg);
ok("keeps BOTH collections (no overwrite)", Object.keys(dtcg).length === 2, "keys=" + JSON.stringify(Object.keys(dtcg)));
ok("local 'primitives' present", "primitives" in dtcg);
ok("global disambiguated to 'primitives (global)'", "primitives (global)" in dtcg);
ok("local token survived (color/bg)", JSON.stringify(dtcg.primitives).includes("$value"));
ok("global token present", JSON.stringify(dtcg["primitives (global)"]).includes("0066ff"));

console.log("\n── Flutter (was: duplicate class name) ──");
const classes = (out.flutter.match(/abstract class (\w+)/g) || []).map(s => s.replace("abstract class ", ""));
ok("two classes emitted", classes.length === 2, "classes=" + JSON.stringify(classes));
ok("class names are unique", new Set(classes).size === classes.length, JSON.stringify(classes));
ok("a global class carries 'Global' suffix", classes.some(c => /Global/.test(c)));

console.log("\n── Figma JSON + CSS (provenance) ──");
const figma = JSON.parse(out.figma);
ok("Figma JSON keeps both (array)", figma.collections.length === 2);
ok("Figma JSON tags global source", figma.collections.some(c => c.source === "global-library" && c.libraryName === "Brand Global"));
ok("CSS marks the global collection", /global — Brand Global/.test(out.css));
ok("CSS emits both vars", out.css.includes("--primitives-color-bg") && out.css.includes("global-primitive"));

console.log("\n── Token search filter ──");
const searched = sandbox.__run(Object.assign({}, data, { search: "skeleton" }));
ok("match count = 1 (only skeleton token)", searched.matchCount === 1, "got " + searched.matchCount);
ok("only collections with matches kept", searched.filteredCols.length === 1 && searched.filteredCols[0] === "primitives", JSON.stringify(searched.filteredCols));
ok("DTCG output contains skeleton", searched.dtcg.includes("skeleton"));
ok("DTCG output drops non-matching token (color/bg)", !searched.dtcg.includes('"bg"'));
ok("CSS output contains only skeleton var", searched.css.includes("skeleton") && !searched.css.includes("--primitives-color-bg"));
const empty = sandbox.__run(Object.assign({}, data, { search: "zzznomatch" }));
ok("no-match search yields 0 tokens", empty.matchCount === 0);
const cleared = sandbox.__run(Object.assign({}, data, { search: "" }));
ok("empty search restores all tokens", cleared.matchCount === 3, "got " + cleared.matchCount);

console.log("\n── Collections collapse ──");
ok("collapse toggles true", sandbox.__collapse(true) === true);
ok("collapse toggles false", sandbox.__collapse(false) === false);

console.log("\n────────────────────────────────────────");
console.log(`  RESULT: ${pass} passed, ${fail} failed`);
console.log("────────────────────────────────────────\n");
process.exit(fail === 0 ? 0 : 1);
