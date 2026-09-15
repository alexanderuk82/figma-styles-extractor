// Realistic data shaped like the real design system: a local primitives /
// semantics pair, a 4-mode responsive collection, aliases across collections,
// and an EXTERNAL library whose collections share names with the local ones.
"use strict";
const color = (hex, r, g, b, a = 1) => ({ type: "color", hex, r, g, b, a });
const num = (v) => ({ type: "number", value: v });
const alias = (name, id, resolved) => ({ type: "alias", aliasName: name, aliasId: id, resolvedValue: resolved });
const same = (modes, v) => Object.fromEntries(modes.map((m) => [m, v]));

const L = ["Light", "Dark"];
const R = ["Mobile Small", "Mobile Large", "Tablet", "Desktop"];

function variables() {
  return {
    _meta: { available: true, fileName: "Acme Tokens", totalVariables: 9, totalCollections: 5 },
    collections: [
      { id: "c-prim", name: "primitives", modes: L.map((m) => ({ id: m, name: m })), variables: [
        { id: "v-blue", name: "color/blue-500", resolvedType: "COLOR", valuesByMode: { Light: color("#0055FF", 0, 85, 255), Dark: color("#3377FF", 51, 119, 255) } },
        { id: "v-white", name: "color/white", resolvedType: "COLOR", valuesByMode: same(L, color("#FFFFFF", 255, 255, 255)) },
        { id: "v-space", name: "space/4", resolvedType: "FLOAT", valuesByMode: same(L, num(16)) },
        { id: "v-alpha", name: "color/overlay", resolvedType: "COLOR", valuesByMode: same(L, color("#000000", 0, 0, 0, 0.4)) },
      ] },
      { id: "c-sem", name: "semantics", modes: L.map((m) => ({ id: m, name: m })), variables: [
        { id: "v-text", name: "text/primary", resolvedType: "COLOR", valuesByMode: { Light: alias("color/blue-500", "v-blue", color("#0055FF", 0, 85, 255)), Dark: alias("color/blue-500", "v-blue", color("#3377FF", 51, 119, 255)) } },
        { id: "v-pad", name: "padding/card", resolvedType: "FLOAT", valuesByMode: same(L, alias("space/4", "v-space", num(16))) },
      ] },
      { id: "c-resp", name: "responsive", modes: R.map((m) => ({ id: m, name: m })), variables: [
        { id: "v-fs", name: "typography/font-size/xl", resolvedType: "FLOAT", valuesByMode: { "Mobile Small": num(24), "Mobile Large": num(28), Tablet: num(32), Desktop: num(36) } },
      ] },
      // external library: same collection names as the local ones — the collision case
      { id: "c-prim-ext", name: "primitives", source: "global-library", remote: true, libraryName: "[Brand] Acme Global (ver2.0)", modes: [{ id: "Mode 1", name: "Mode 1" }], variables: [
        { id: "v-ext-blue", name: "color/blue-500", resolvedType: "COLOR", valuesByMode: { "Mode 1": color("#0066FF", 0, 102, 255) } },
        { id: "v-ext-brand", name: "brand-primitive/type/font-size/l", resolvedType: "FLOAT", valuesByMode: { "Mode 1": num(20) } },
      ] },
      { id: "c-sem-ext", name: "semantics", source: "global-library", remote: true, libraryName: "[Brand] Acme Global (ver2.0)", modes: [{ id: "Mode 1", name: "Mode 1" }], variables: [
        { id: "v-ext-text", name: "text/brand", resolvedType: "COLOR", valuesByMode: { "Mode 1": alias("color/blue-500", "v-ext-blue", color("#0066FF", 0, 102, 255)) } },
      ] },
    ],
  };
}
const allModes = () => [...L, ...R, "Mode 1"];

function styles() {
  return {
    _meta: { fileName: "Acme Tokens", exportedAt: "2026-09-15T10:00:00.000Z", totalStyles: 4 },
    paintStyles: [
      { name: "brand/red", description: "Primary", paints: [{ type: "SOLID", color: { r: 255, g: 0, b: 0, a: 1, hex: "#FF0000" } }] },
      { name: "brand/fade", description: "", paints: [{ type: "GRADIENT_LINEAR", gradientTransform: [[1, 0, 0], [0, 1, 0]], gradientStops: [{ position: 0, color: { r: 255, g: 0, b: 0, a: 1, hex: "#FF0000" } }, { position: 1, color: { r: 0, g: 0, b: 255, a: 1, hex: "#0000FF" } }] }] },
    ],
    textStyles: [
      { name: "body/regular", description: "", fontFamily: "Inter", fontStyle: "Regular", fontSize: 16, lineHeight: { unit: "PERCENT", value: 150 }, letterSpacing: { value: 0, unit: "PIXELS" }, paragraphSpacing: 0, boundVariables: {} },
      { name: "heading/xl", description: "", fontFamily: "Inter", fontStyle: "Bold", fontSize: 32, lineHeight: { unit: "AUTO" }, letterSpacing: { value: -0.5, unit: "PIXELS" }, paragraphSpacing: 0, boundVariables: { fontSize: "typography/font-size/xl" } },
    ],
    effectStyles: [
      { name: "shadow/card", description: "", effects: [{ type: "DROP_SHADOW", visible: true, color: { r: 0, g: 0, b: 0, a: 0.2, hex: "#000000" }, offset: { x: 0, y: 4 }, radius: 12, spread: 0, blendMode: "NORMAL" }] },
    ],
    gridStyles: [],
  };
}

module.exports = { variables, styles, allModes, L, R };
