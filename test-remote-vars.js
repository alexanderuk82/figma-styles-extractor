// Internal test harness for the GLOBAL / library variable discovery layer.
// Extracts the REMOTE-VARS region from code.js verbatim and runs it against a
// mocked Team Library API: an EXTERNAL global library (primitives + responsive
// with cross-collection alias chains) plus the file's OWN published library
// (which must be skipped as it duplicates the local export).
// Run with: node test-remote-vars.js
"use strict";
const fs = require("fs");
const path = require("path");

// ── Extract the testable region verbatim from code.js ──
const src = fs.readFileSync(path.join(__dirname, "code.js"), "utf8");
const begin = src.indexOf("// ===== REMOTE-VARS-BEGIN");
const end = src.indexOf("// ===== REMOTE-VARS-END");
if (begin === -1 || end === -1) { console.error("✗ region markers not found"); process.exit(1); }
const region = src.slice(begin, end);

function rgbToHex(r, g, b) {
  const toHex = (v) => { const h = Math.round(v * 255).toString(16); return h.length === 1 ? "0" + h : h; };
  return "#" + toHex(r) + toHex(g) + toHex(b);
}

// ── Mock world ──
const SELF_NAME = "[Design System] My File";
const COL_PRIM = { id: "VC:prim", name: "primitives", defaultModeId: "M1", modes: [{ modeId: "M1", name: "Mode 1" }] };
const COL_RESP = { id: "VC:resp", name: "responsive", defaultModeId: "L", modes: [{ modeId: "L", name: "Light" }, { modeId: "D", name: "Dark" }] };
const COL_SELF = { id: "VC:self", name: "components", defaultModeId: "S1", modes: [{ modeId: "S1", name: "Value" }] };

// Variables keyed by id and by key
const VARS_BY_ID = {
  "V:brand": { id: "V:brand", key: "k_brand", name: "global-primitive/color/brand", resolvedType: "COLOR", variableCollectionId: "VC:prim", description: "", scopes: [], codeSyntax: {}, valuesByMode: { M1: { r: 0, g: 0.4, b: 1, a: 1 } } },
  "V:base":  { id: "V:base",  key: "k_base",  name: "global-primitive/size/base",  resolvedType: "FLOAT", variableCollectionId: "VC:prim", description: "", scopes: [], codeSyntax: {}, valuesByMode: { M1: 8 } },
  "V:gap":   { id: "V:gap",   key: "k_gap",   name: "global-responsive/gap/sm",    resolvedType: "FLOAT", variableCollectionId: "VC:resp", description: "", scopes: [], codeSyntax: {}, valuesByMode: { L: { type: "VARIABLE_ALIAS", id: "V:base" }, D: { type: "VARIABLE_ALIAS", id: "V:base" } } },
  "V:self":  { id: "V:self",  key: "k_self",  name: "components/button/bg",        resolvedType: "COLOR", variableCollectionId: "VC:self", description: "", scopes: [], codeSyntax: {}, valuesByMode: { S1: { r: 1, g: 0, b: 0, a: 1 } } },
};
const VARS_BY_KEY = {};
for (const id in VARS_BY_ID) VARS_BY_KEY[VARS_BY_ID[id].key] = VARS_BY_ID[id];
const COLS = { "VC:prim": COL_PRIM, "VC:resp": COL_RESP, "VC:self": COL_SELF };

// Library collections: 2 external (Brand Global) + 1 self-published (must be skipped)
const LIB_COLS = [
  { key: "lk_prim", name: "primitives",  libraryName: "Brand Global" },
  { key: "lk_resp", name: "responsive",  libraryName: "Brand Global" },
  { key: "lk_self", name: "components",  libraryName: SELF_NAME },
];
const LIB_VARS = {
  "lk_prim": [{ key: "k_brand", name: "global-primitive/color/brand", resolvedType: "COLOR" }, { key: "k_base", name: "global-primitive/size/base", resolvedType: "FLOAT" }],
  "lk_resp": [{ key: "k_gap", name: "global-responsive/gap/sm", resolvedType: "FLOAT" }],
  "lk_self": [{ key: "k_self", name: "components/button/bg", resolvedType: "COLOR" }],
};

let importCount = 0;
const figma = {
  root: { name: SELF_NAME },
  teamLibrary: {
    getAvailableLibraryVariableCollectionsAsync: async () => LIB_COLS,
    getVariablesInLibraryCollectionAsync: async (key) => LIB_VARS[key] || [],
  },
  variables: {
    importVariableByKeyAsync: async (key) => { importCount++; if (!VARS_BY_KEY[key]) throw new Error("bad key"); return VARS_BY_KEY[key]; },
    getVariableByIdAsync: async (id) => VARS_BY_ID[id] || null,
    getVariableCollectionByIdAsync: async (id) => COLS[id] || null,
  },
};

const factory = new Function("figma", "rgbToHex",
  region + "\nreturn { resolveVariableValueAsync, resolveAliasChainAsync, importLibraryVariablesParallel, extractLibraryVariables };");
const M = factory(figma, rgbToHex);

let pass = 0, fail = 0;
function ok(n, c, e) { if (c) { pass++; console.log("  ✅ " + n); } else { fail++; console.log("  ❌ " + n + (e ? "  → " + e : "")); } }
function eq(n, a, b) { ok(n, JSON.stringify(a) === JSON.stringify(b), "got " + JSON.stringify(a) + ", want " + JSON.stringify(b)); }

(async () => {
  console.log("\n── 1. importLibraryVariablesParallel imports in batches ──");
  const imp = await M.importLibraryVariablesParallel([
    { key: "k_brand", libraryName: "Brand Global" },
    { key: "k_base", libraryName: "Brand Global" },
    { key: "k_gap", libraryName: "Brand Global" },
  ], 2);
  eq("imported 3", imp.length, 3);
  ok("carries libraryName", imp.every((x) => x.libraryName === "Brand Global"));
  ok("carries variable objects", imp.every((x) => x.variable && x.variable.id));

  console.log("\n── 2. resolveVariableValueAsync: direct colour + cross-collection alias ──");
  const brand = await M.resolveVariableValueAsync(VARS_BY_ID["V:brand"], "M1", "Mode 1");
  eq("brand colour hex", brand.hex, rgbToHex(0, 0.4, 1));
  const gap = await M.resolveVariableValueAsync(VARS_BY_ID["V:gap"], "L", "Light");
  eq("gap is alias", gap.type, "alias");
  eq("gap alias name", gap.aliasName, "global-primitive/size/base");
  eq("gap resolved value", gap.resolvedValue, { type: "number", value: 8 });

  console.log("\n── 3. extractLibraryVariables → external globals only, grouped, tagged ──");
  importCount = 0;
  const cols = await M.extractLibraryVariables();
  eq("two external collections (self skipped)", cols.length, 2);
  const names = cols.map((c) => c.name).sort();
  eq("collection names", names, ["primitives", "responsive"]);
  ok("self 'components' collection skipped", !cols.some((c) => c.libraryName === SELF_NAME));
  ok("every collection tagged source=global-library", cols.every((c) => c.source === "global-library"));
  ok("every collection tagged libraryName=Brand Global", cols.every((c) => c.libraryName === "Brand Global"));
  ok("every collection flagged remote", cols.every((c) => c.remote === true));
  const prim = cols.find((c) => c.name === "primitives");
  const resp = cols.find((c) => c.name === "responsive");
  eq("primitives var count", prim.variableCount, 2);
  eq("responsive var count", resp.variableCount, 1);
  eq("responsive modes preserved", resp.modes.map((m) => m.name), ["Light", "Dark"]);
  ok("variables tagged source+libraryName", prim.variables.every((v) => v.source === "global-library" && v.libraryName === "Brand Global" && v.remote === true));
  const brandVar = prim.variables.find((v) => v.name === "global-primitive/color/brand");
  eq("brand resolved in catalogue", brandVar.valuesByMode["Mode 1"].hex, rgbToHex(0, 0.4, 1));
  const gapVar = resp.variables.find((v) => v.name === "global-responsive/gap/sm");
  eq("gap alias resolved per mode", gapVar.valuesByMode.Light.resolvedValue, { type: "number", value: 8 });
  ok("did NOT import the self-library variable", importCount === 3, "importCount=" + importCount);

  console.log("\n── 4. Graceful when Team Library API absent ──");
  const f2 = new Function("figma", "rgbToHex", region + "\nreturn { extractLibraryVariables };");
  const M2 = f2({ teamLibrary: undefined, variables: {}, root: { name: "x" } }, rgbToHex);
  eq("returns [] when teamLibrary missing", await M2.extractLibraryVariables(), []);

  console.log("\n────────────────────────────────────────");
  console.log(`  RESULT: ${pass} passed, ${fail} failed`);
  console.log("────────────────────────────────────────\n");
  process.exit(fail === 0 ? 0 : 1);
})();
