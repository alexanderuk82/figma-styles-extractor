// DS Styles & Variables Extractor v2.0
// Extracts Styles (Paint, Text, Effect, Grid)
// Extracts Variables (all Collections, all Modes, resolved values)
// Supports live SYNC + multi-format output (JSON, CSS, Flutter, W3C DTCG)

figma.showUI(__html__, { width: 900, height: 860, themeColors: true });

// ─── Colour Helpers ───

function rgbToHex(r, g, b) {
  const toHex = (v) => {
    const h = Math.round(v * 255).toString(16);
    return h.length === 1 ? "0" + h : h;
  };
  return "#" + toHex(r) + toHex(g) + toHex(b);
}

// ─── Style Extractors ───

function extractPaint(paint) {
  const base = {
    type: paint.type,
    visible: paint.visible !== false,
    opacity: paint.opacity !== undefined ? paint.opacity : 1,
    blendMode: paint.blendMode || "NORMAL",
  };
  if (paint.type === "SOLID") {
    base.color = {
      hex: rgbToHex(paint.color.r, paint.color.g, paint.color.b),
      r: Math.round(paint.color.r * 255),
      g: Math.round(paint.color.g * 255),
      b: Math.round(paint.color.b * 255),
      a: paint.opacity !== undefined ? paint.opacity : 1,
    };
  }
  if (paint.type.startsWith("GRADIENT_")) {
    base.gradientStops = paint.gradientStops.map((stop) => ({
      position: stop.position,
      color: {
        hex: rgbToHex(stop.color.r, stop.color.g, stop.color.b),
        r: Math.round(stop.color.r * 255),
        g: Math.round(stop.color.g * 255),
        b: Math.round(stop.color.b * 255),
        a: stop.color.a,
      },
    }));
    if (paint.gradientTransform) base.gradientTransform = paint.gradientTransform;
  }
  if (paint.type === "IMAGE") {
    base.scaleMode = paint.scaleMode;
    base.imageHash = paint.imageHash;
  }
  return base;
}

function extractEffect(effect) {
  const base = { type: effect.type, visible: effect.visible !== false };
  if (effect.color) {
    base.color = {
      hex: rgbToHex(effect.color.r, effect.color.g, effect.color.b),
      r: Math.round(effect.color.r * 255),
      g: Math.round(effect.color.g * 255),
      b: Math.round(effect.color.b * 255),
      a: effect.color.a,
    };
  }
  if (effect.offset) base.offset = { x: effect.offset.x, y: effect.offset.y };
  if (effect.radius !== undefined) base.radius = effect.radius;
  if (effect.spread !== undefined) base.spread = effect.spread;
  if (effect.blendMode) base.blendMode = effect.blendMode;
  return base;
}

function extractLineHeight(lh) {
  if (!lh || lh.unit === "AUTO") return { unit: "AUTO" };
  return { value: lh.value, unit: lh.unit };
}

function extractLetterSpacing(ls) {
  if (!ls) return { value: 0, unit: "PIXELS" };
  return { value: ls.value, unit: ls.unit };
}

// ─── Extract all Styles ───

function extractAllStyles() {
  const result = {
    _meta: {
      exportedAt: new Date().toISOString(),
      fileName: figma.root.name,
      totalStyles: 0,
    },
    paintStyles: [],
    textStyles: [],
    effectStyles: [],
    gridStyles: [],
  };

  for (const s of figma.getLocalPaintStyles()) {
    result.paintStyles.push({
      name: s.name,
      description: s.description || "",
      paints: s.paints.map(extractPaint),
    });
  }

  for (const s of figma.getLocalTextStyles()) {
    // Extract bound variable names for text style properties
    var boundVars = {};
    // Resolve bound variable values for ALL modes of their collection
    var boundVarModes = {};
    var boundVarModeNames = null; // modes from the first bound variable's collection
    var boundVarValues = {}; // props bound to single-mode collections: one resolved raw value
    if (s.boundVariables) {
      var bvKeys = Object.keys(s.boundVariables);
      for (var bk = 0; bk < bvKeys.length; bk++) {
        var prop = bvKeys[bk];
        var binding = s.boundVariables[prop];
        if (binding && binding.id) {
          try {
            var bVar = figma.variables.getVariableById(binding.id);
            if (bVar) {
              boundVars[prop] = bVar.name;
              // Resolve this variable for all modes of its collection
              var col = figma.variables.getVariableCollectionById(bVar.variableCollectionId);
              if (col && col.modes.length > 1) {
                boundVarModes[prop] = {};
                for (var mi = 0; mi < col.modes.length; mi++) {
                  var mode = col.modes[mi];
                  var resolved = resolveVariableValue(bVar, mode.modeId, mode.name);
                  if (resolved) {
                    // Preserve alias chain so docs can show alias name (matches Figma's responsive collection view)
                    if (resolved.type === 'alias') {
                      boundVarModes[prop][mode.name] = {
                        type: 'alias',
                        aliasName: resolved.aliasName,
                        resolvedValue: resolved.resolvedValue,
                      };
                    } else {
                      boundVarModes[prop][mode.name] = resolved;
                    }
                  }
                }
                // Capture mode names from the first multi-mode collection
                if (!boundVarModeNames) {
                  boundVarModeNames = col.modes.map(function(m) { return m.name; });
                }
              } else if (col && col.modes.length === 1) {
                // Single-mode collection (e.g. a primitives library) — resolve its
                // only mode so docs can show the raw value under the token name
                var onlyMode = col.modes[0];
                var resolvedSingle = resolveVariableValue(bVar, onlyMode.modeId, onlyMode.name);
                if (resolvedSingle) boundVarValues[prop] = resolvedSingle;
              }
            }
          } catch (e) {}
        }
      }
    }
    var textStyleEntry = {
      name: s.name,
      description: s.description || "",
      fontFamily: s.fontName.family,
      fontStyle: s.fontName.style,
      fontSize: s.fontSize,
      lineHeight: extractLineHeight(s.lineHeight),
      letterSpacing: extractLetterSpacing(s.letterSpacing),
      textCase: s.textCase || "ORIGINAL",
      textDecoration: s.textDecoration || "NONE",
      paragraphSpacing: s.paragraphSpacing || 0,
      paragraphIndent: s.paragraphIndent || 0,
      boundVariables: boundVars,
    };
    // Include per-mode resolved values if any bound variables have multiple modes
    if (Object.keys(boundVarModes).length > 0) {
      textStyleEntry.boundVarModes = boundVarModes;
      if (boundVarModeNames) textStyleEntry.modeNames = boundVarModeNames;
    }
    // Include resolved values for props bound to single-mode collections
    if (Object.keys(boundVarValues).length > 0) {
      textStyleEntry.boundVarValues = boundVarValues;
    }
    result.textStyles.push(textStyleEntry);
  }

  for (const s of figma.getLocalEffectStyles()) {
    result.effectStyles.push({
      name: s.name,
      description: s.description || "",
      effects: s.effects.map(extractEffect),
    });
  }

  for (const s of figma.getLocalGridStyles()) {
    result.gridStyles.push({
      name: s.name,
      description: s.description || "",
      grids: s.layoutGrids.map((g) => ({
        pattern: g.pattern,
        sectionSize: g.sectionSize,
        visible: g.visible,
        color: g.color
          ? { hex: rgbToHex(g.color.r, g.color.g, g.color.b), a: g.color.a }
          : undefined,
        alignment: g.alignment,
        gutterSize: g.gutterSize,
        offset: g.offset,
        count: g.count,
      })),
    });
  }

  result._meta.totalStyles =
    result.paintStyles.length +
    result.textStyles.length +
    result.effectStyles.length +
    result.gridStyles.length;

  result._meta.breakdown = {
    paint: result.paintStyles.length,
    text: result.textStyles.length,
    effect: result.effectStyles.length,
    grid: result.gridStyles.length,
  };

  return result;
}

// ─── Variable Value Resolver ───

function resolveAliasChain(aliasVar, modeName, depth) {
  if (!depth) depth = 0;
  if (depth > 10) return null;

  var col;
  try { col = figma.variables.getVariableCollectionById(aliasVar.variableCollectionId); } catch(e) { return null; }
  if (!col) return null;

  // Find matching mode by name, fallback to default mode
  var modeId = col.defaultModeId;
  for (var m = 0; m < col.modes.length; m++) {
    if (col.modes[m].name === modeName) {
      modeId = col.modes[m].modeId;
      break;
    }
  }

  var value = aliasVar.valuesByMode[modeId];
  if (value === undefined || value === null) return null;

  // If still an alias, follow chain
  if (typeof value === "object" && value.type === "VARIABLE_ALIAS") {
    try {
      var next = figma.variables.getVariableById(value.id);
      if (next) return resolveAliasChain(next, modeName, depth + 1);
    } catch(e) {}
    return null;
  }

  // Concrete value
  if (aliasVar.resolvedType === "COLOR" && typeof value === "object") {
    return {
      type: "color",
      hex: rgbToHex(value.r, value.g, value.b),
      r: Math.round(value.r * 255),
      g: Math.round(value.g * 255),
      b: Math.round(value.b * 255),
      a: value.a !== undefined ? +value.a.toFixed(4) : 1,
    };
  }
  if (aliasVar.resolvedType === "FLOAT") return { type: "number", value: value };
  if (aliasVar.resolvedType === "STRING") return { type: "string", value: value };
  if (aliasVar.resolvedType === "BOOLEAN") return { type: "boolean", value: value };
  return null;
}

function resolveVariableValue(variable, modeId, modeName) {
  const value = variable.valuesByMode[modeId];
  if (value === undefined || value === null) return null;

  // Handle alias (variable referencing another variable)
  if (typeof value === "object" && value.type === "VARIABLE_ALIAS") {
    try {
      const aliasVar = figma.variables.getVariableById(value.id);
      if (aliasVar) {
        var resolved = modeName ? resolveAliasChain(aliasVar, modeName, 0) : null;
        return {
          type: "alias",
          aliasName: aliasVar.name,
          aliasId: aliasVar.id,
          resolvedValue: resolved,
        };
      }
    } catch (e) {
      return { type: "alias", aliasName: "unresolved", aliasId: value.id };
    }
  }

  // COLOR
  if (variable.resolvedType === "COLOR" && typeof value === "object") {
    return {
      type: "color",
      hex: rgbToHex(value.r, value.g, value.b),
      r: Math.round(value.r * 255),
      g: Math.round(value.g * 255),
      b: Math.round(value.b * 255),
      a: value.a !== undefined ? +value.a.toFixed(4) : 1,
    };
  }

  // FLOAT
  if (variable.resolvedType === "FLOAT") {
    return { type: "number", value: value };
  }

  // STRING
  if (variable.resolvedType === "STRING") {
    return { type: "string", value: value };
  }

  // BOOLEAN
  if (variable.resolvedType === "BOOLEAN") {
    return { type: "boolean", value: value };
  }

  return { type: "unknown", value: String(value) };
}

// ─── Extract all Variables ───

function extractAllVariables() {
  // Check if Variables API is available
  if (!figma.variables || !figma.variables.getLocalVariableCollections) {
    return {
      _meta: { available: false, reason: "Variables API not available" },
      collections: [],
    };
  }

  const collections = figma.variables.getLocalVariableCollections();
  const allVariables = figma.variables.getLocalVariables();

  const result = {
    _meta: {
      exportedAt: new Date().toISOString(),
      fileName: figma.root.name,
      totalCollections: collections.length,
      totalVariables: allVariables.length,
      available: true,
    },
    collections: [],
  };

  for (const col of collections) {
    const colVars = allVariables.filter(
      (v) => v.variableCollectionId === col.id
    );

    // Group variables by folder path (slash-separated names)
    const variables = colVars.map((v) => {
      const valuesByMode = {};
      for (const mode of col.modes) {
        valuesByMode[mode.name] = resolveVariableValue(v, mode.modeId, mode.name);
      }

      return {
        id: v.id,
        name: v.name,
        resolvedType: v.resolvedType,
        description: v.description || "",
        scopes: v.scopes || [],
        codeSyntax: v.codeSyntax || {},
        valuesByMode: valuesByMode,
      };
    });

    result.collections.push({
      id: col.id,
      name: col.name,
      modes: col.modes.map((m) => ({ id: m.modeId, name: m.name })),
      variableCount: variables.length,
      variables: variables,
    });
  }

  return result;
}

// ─── Global / Library Variable Discovery ──────────────────────────────
// Additive layer: the local export above (extractAllVariables) is unchanged
// and still LOCAL-ONLY. This layer pulls the entire GLOBAL library catalogue
// (every collection/variable published by the libraries enabled in this file —
// used or not) via the Team Library API, so globals that components wire to
// directly — with no local intermediary — are recognised on export. Each item
// is tagged source:"global-library" + libraryName and grouped by collection so
// the UI and the git export can mark what comes from Global. Fully wrapped in
// try/catch so a failure here can never break the existing local export.
// ===== REMOTE-VARS-BEGIN (testable region — keep self-contained) =====

// Per-run caches so alias chains don't re-fetch the same primitives/collections
// (a global responsive token can alias-resolve through dozens of shared
// primitives). Reset at the start of extractLibraryVariables.
var _rvVarCache = {};
var _rvColCache = {};
async function _rvGetVar(id) {
  if (Object.prototype.hasOwnProperty.call(_rvVarCache, id)) return _rvVarCache[id];
  var v = null;
  try { v = await figma.variables.getVariableByIdAsync(id); } catch (e) { v = null; }
  _rvVarCache[id] = v;
  return v;
}
async function _rvGetCol(id) {
  if (Object.prototype.hasOwnProperty.call(_rvColCache, id)) return _rvColCache[id];
  var c = null;
  try { c = await figma.variables.getVariableCollectionByIdAsync(id); } catch (e) { c = null; }
  _rvColCache[id] = c;
  return c;
}

// Async, remote-aware mirror of resolveAliasChain (uses *Async lookups so
// it can follow alias chains that live in remote/library collections).
async function resolveAliasChainAsync(aliasVar, modeName, depth) {
  if (!depth) depth = 0;
  if (depth > 10) return null;

  var col = await _rvGetCol(aliasVar.variableCollectionId);
  if (!col) return null;

  var modeId = col.defaultModeId;
  for (var m = 0; m < col.modes.length; m++) {
    if (col.modes[m].name === modeName) { modeId = col.modes[m].modeId; break; }
  }

  var value = aliasVar.valuesByMode[modeId];
  if (value === undefined || value === null) return null;

  if (typeof value === "object" && value.type === "VARIABLE_ALIAS") {
    var next = await _rvGetVar(value.id);
    if (next) return await resolveAliasChainAsync(next, modeName, depth + 1);
    return null;
  }

  if (aliasVar.resolvedType === "COLOR" && typeof value === "object") {
    return {
      type: "color",
      hex: rgbToHex(value.r, value.g, value.b),
      r: Math.round(value.r * 255),
      g: Math.round(value.g * 255),
      b: Math.round(value.b * 255),
      a: value.a !== undefined ? +value.a.toFixed(4) : 1,
    };
  }
  if (aliasVar.resolvedType === "FLOAT") return { type: "number", value: value };
  if (aliasVar.resolvedType === "STRING") return { type: "string", value: value };
  if (aliasVar.resolvedType === "BOOLEAN") return { type: "boolean", value: value };
  return null;
}

// Async, remote-aware mirror of resolveVariableValue.
async function resolveVariableValueAsync(variable, modeId, modeName) {
  var value = variable.valuesByMode[modeId];
  if (value === undefined || value === null) return null;

  if (typeof value === "object" && value.type === "VARIABLE_ALIAS") {
    var aliasVar = await _rvGetVar(value.id);
    if (aliasVar) {
      var resolved = modeName ? await resolveAliasChainAsync(aliasVar, modeName, 0) : null;
      return { type: "alias", aliasName: aliasVar.name, aliasId: aliasVar.id, resolvedValue: resolved };
    }
    return { type: "alias", aliasName: "unresolved", aliasId: value.id };
  }

  if (variable.resolvedType === "COLOR" && typeof value === "object") {
    return {
      type: "color",
      hex: rgbToHex(value.r, value.g, value.b),
      r: Math.round(value.r * 255),
      g: Math.round(value.g * 255),
      b: Math.round(value.b * 255),
      a: value.a !== undefined ? +value.a.toFixed(4) : 1,
    };
  }
  if (variable.resolvedType === "FLOAT") return { type: "number", value: value };
  if (variable.resolvedType === "STRING") return { type: "string", value: value };
  if (variable.resolvedType === "BOOLEAN") return { type: "boolean", value: value };
  return { type: "unknown", value: String(value) };
}

// Import a list of library variable keys, in parallel batches. Sequential
// importVariableByKeyAsync is ~165ms/var (≈90s for 537 globals); batches of 25
// overlap the round-trips and bring the same set down to ≈8s. Returns an array
// of { variable, libraryName } for everything that imported successfully.
async function importLibraryVariablesParallel(entries, batchSize) {
  if (!batchSize) batchSize = 25;
  var imported = [];
  for (var b = 0; b < entries.length; b += batchSize) {
    var slice = entries.slice(b, b + batchSize);
    var results = await Promise.all(slice.map(function (en) {
      return figma.variables.importVariableByKeyAsync(en.key)
        .then(function (v) { return v ? { variable: v, libraryName: en.libraryName } : null; })
        .catch(function () { return null; });
    }));
    for (var ri = 0; ri < results.length; ri++) {
      if (results[ri]) imported.push(results[ri]);
    }
  }
  return imported;
}

// Orchestrator: pull the entire GLOBAL library catalogue via the Team Library
// API (no node traversal — that timed out at >30s on this 81-page file). Lists
// every collection/variable published by the libraries ENABLED in this file,
// imports each (in parallel) to read full values/modes/alias chains, and groups
// them by collection. Collections published by THIS file are skipped — they are
// already covered by the local export (extractAllVariables). Every returned
// collection and variable is tagged source:"global-library" + libraryName so
// the UI and the Bitbucket/GitHub export can clearly mark what comes from Global.
async function extractLibraryVariables() {
  if (!figma.teamLibrary || !figma.teamLibrary.getAvailableLibraryVariableCollectionsAsync) return [];
  if (!figma.variables || !figma.variables.importVariableByKeyAsync) return [];

  // Reset per-run alias caches
  _rvVarCache = {};
  _rvColCache = {};

  var libCols;
  try { libCols = await figma.teamLibrary.getAvailableLibraryVariableCollectionsAsync(); } catch (e) { return []; }
  if (!libCols || !libCols.length) return [];

  var selfName = (figma.root && figma.root.name) || "";

  // Enumerate every variable key in EXTERNAL libraries only (skip this file's
  // own published collections — they duplicate the local export).
  var entries = [];
  for (var ci = 0; ci < libCols.length; ci++) {
    var lc = libCols[ci];
    if (lc.libraryName === selfName) continue;
    var libVars = [];
    try { libVars = await figma.teamLibrary.getVariablesInLibraryCollectionAsync(lc.key); } catch (e) { libVars = []; }
    for (var vi = 0; vi < libVars.length; vi++) {
      entries.push({ key: libVars[vi].key, libraryName: lc.libraryName });
    }
  }
  if (!entries.length) return [];

  // Safety cap (logged, never silent) so a pathologically large library can't hang the export.
  var MAX = 4000;
  if (entries.length > MAX) {
    console.warn("[LibraryVars] " + entries.length + " library variables; importing first " + MAX + " (cap). Some globals omitted.");
    entries = entries.slice(0, MAX);
  }

  var imported = await importLibraryVariablesParallel(entries, 25);

  // Every imported variable is already in hand: seed the cache so alias chains
  // resolve from memory instead of making another API round-trip per hop.
  for (var pc = 0; pc < imported.length; pc++) {
    var pv = imported[pc].variable;
    if (pv && pv.id) _rvVarCache[pv.id] = pv;
  }

  // Group by the imported variable's real (library) collection
  var byCollection = {};
  var order = [];
  for (var iv = 0; iv < imported.length; iv++) {
    var v = imported[iv].variable;
    var col = await _rvGetCol(v.variableCollectionId);
    if (!col) continue;
    if (!byCollection[col.id]) {
      byCollection[col.id] = { collection: col, libraryName: imported[iv].libraryName || "", vars: [] };
      order.push(col.id);
    }
    byCollection[col.id].vars.push(v);
  }

  var out = [];
  for (var o = 0; o < order.length; o++) {
    var bucket = byCollection[order[o]];
    var bcol = bucket.collection;
    // Resolving one variable-and-mode at a time was the bottleneck: a library of a
    // few thousand variables across four modes meant thousands of sequential awaits,
    // which is what made the plugin take minutes to open. Resolve in chunks instead.
    var variables = [];
    var CHUNK = 40;
    for (var x = 0; x < bucket.vars.length; x += CHUNK) {
      var slice = bucket.vars.slice(x, x + CHUNK);
      var resolved = await Promise.all(slice.map(function (bv) {
        var valuesByMode = {};
        return Promise.all(bcol.modes.map(function (m) {
          return resolveVariableValueAsync(bv, m.modeId, m.name).then(function (val) { valuesByMode[m.name] = val; });
        })).then(function () {
          return {
            id: bv.id,
            name: bv.name,
            resolvedType: bv.resolvedType,
            description: bv.description || "",
            scopes: bv.scopes || [],
            codeSyntax: bv.codeSyntax || {},
            valuesByMode: valuesByMode,
            remote: true,
            source: "global-library",
            libraryName: bucket.libraryName,
          };
        });
      }));
      for (var rr = 0; rr < resolved.length; rr++) variables.push(resolved[rr]);
    }
    out.push({
      id: bcol.id,
      name: bcol.name,
      remote: true,
      source: "global-library",
      libraryName: bucket.libraryName,
      modes: bcol.modes.map(function (m) { return { id: m.modeId, name: m.name }; }),
      variableCount: variables.length,
      variables: variables,
    });
  }
  return out;
}
// ===== REMOTE-VARS-END =====

// ─── Send data to UI ───

async function sendAllData() {
  const styles = extractAllStyles();
  const variables = extractAllVariables();

  // Send the local data straight away. The global library can take a while on a
  // big design system, and there is no reason to keep the whole panel blank
  // while it loads — it arrives as a second update.
  figma.ui.postMessage({
    type: "all-data",
    payload: { styles, variables },
    globalsPending: true,
  });

  // Additive: pull the full GLOBAL library catalogue and append it, clearly
  // tagged (source:"global-library" + libraryName) and grouped by collection.
  try {
    var hasTeamLib = !!(figma.teamLibrary && figma.teamLibrary.getAvailableLibraryVariableCollectionsAsync);
    var globalCollections = await extractLibraryVariables();
    if (globalCollections && globalCollections.length && variables.collections) {
      for (var rc = 0; rc < globalCollections.length; rc++) {
        variables.collections.push(globalCollections[rc]);
      }
      var gVars = globalCollections.reduce(function (n, col) { return n + col.variables.length; }, 0);
      if (variables._meta) {
        variables._meta.totalCollections = variables.collections.length;
        var tv = 0;
        for (var c = 0; c < variables.collections.length; c++) tv += variables.collections[c].variables.length;
        variables._meta.totalVariables = tv;
        variables._meta.globalCollections = globalCollections.length;
        variables._meta.globalVariables = gVars;
      }
      figma.ui.postMessage({
        type: "all-data",
        payload: { styles, variables },
        globalsPending: false,
        globalsLoaded: { collections: globalCollections.length, variables: gVars },
      });
    } else {
      figma.ui.postMessage({ type: "globals-done", ok: false, reason: hasTeamLib ? "no external library" : "teamLibrary unavailable" });
    }
  } catch (e) {
    console.error("[GlobalVars] library import failed (local export unaffected):", e);
    figma.ui.postMessage({ type: "globals-done", ok: false, reason: (e && e.message) || String(e) });
  }
}

// ─── Documentation Generator ───

const DC = {
  W: 1400, P: 48, EX: 240, GAP: 32, MIN: 16,
  headerBg:  { r: 0.118, g: 0.161, b: 0.231 },
  white:     { r: 1, g: 1, b: 1 },
  text:      { r: 0.118, g: 0.161, b: 0.231 },
  textSec:   { r: 0.392, g: 0.455, b: 0.545 },
  badgeBg:   { r: 0.945, g: 0.941, b: 0.933 },
  badgeTxt:  { r: 0.278, g: 0.333, b: 0.412 },
  divider:   { r: 0.886, g: 0.910, b: 0.941 },
  rawValue:  { r: 0.761, g: 0.094, b: 0.357 }, // #C2185B — 5.9:1 on white, WCAG AA at 16px
  rawChipBg: { r: 0.988, g: 0.906, b: 0.937 }, // #FCE7EF — with #C2185B text: 5.0:1, WCAG AA at 16px
  descBg:    { r: 0.910, g: 0.945, b: 0.984 }, // #E8F1FB — with DC.text: 12.8:1, WCAG AAA at 16px
  // Audit statuses — all verified at 16px: amber+near-black 8.1:1 (AAA),
  // red/green/purple with white text 5.0:1, 5.4:1 and 5.7:1 (AA)
  auOk:      { r: 0.082, g: 0.478, b: 0.267 }, // #157A44
  auDep:     { r: 0.878, g: 0.659, b: 0.000 }, // #E0A800
  auDepTxt:  { r: 0.102, g: 0.102, b: 0.102 }, // #1A1A1A on amber
  auMiss:    { r: 0.769, g: 0.204, b: 0.106 }, // #C4341B
  auInfo:    { r: 0.420, g: 0.310, b: 0.839 }, // #6B4FD6
  auComp:    { r: 0.639, g: 0.310, b: 0.000 }, // #A34F00 — white text 5.7:1 (AA)
  // Health score on the dark header (#1E293B): 7.8:1, 8.2:1 and 5.7:1 — AAA, AAA, AA
  hpGood:    { r: 0.310, g: 0.831, b: 0.541 }, // #4FD48A
  hpWarn:    { r: 0.949, g: 0.722, b: 0.294 }, // #F2B84B
  hpBad:     { r: 1.000, g: 0.478, b: 0.361 }, // #FF7A5C
};

const _lf = new Set();

// ─── Live Sync State ───
var docState = {
  active: false,
  frames: [],
  rows: new Map(),
  snapshot: { paint: new Map(), text: new Map(), effect: new Map(), variable: new Map() },
  varModes: [],
};
var _liveTimer = null;
var _isLiveUpdating = false;
var _varPollTimer = null;
var _varPollBusy = false;

function startVarPolling() {
  if (_varPollTimer) return;
  _varPollTimer = setInterval(function () {
    if (!docState.active || docState.snapshot.variable.size === 0) {
      stopVarPolling();
      return;
    }
    performVariablePoll();
  }, 3000);
  console.log("[VarPoll] Started polling every 3s");
}

function stopVarPolling() {
  if (_varPollTimer) {
    clearInterval(_varPollTimer);
    _varPollTimer = null;
    _varPollBusy = false;
    console.log("[VarPoll] Stopped polling");
  }
}

// Lightweight variable poll — own flag, independent from style sync
async function performVariablePoll() {
  if (_varPollBusy || !docState.active || docState.snapshot.variable.size === 0) return;
  _varPollBusy = true;

  try {
    // Check variable doc frames still exist
    var anyVarAlive = false;
    for (var fi = 0; fi < docState.frames.length; fi++) {
      if (docState.frames[fi].isVariable && figma.getNodeById(docState.frames[fi].frameId)) {
        anyVarAlive = true; break;
      }
    }
    if (!anyVarAlive) {
      var vk = [];
      docState.rows.forEach(function(v, k) { if (k.startsWith("var:")) vk.push(k); });
      for (var vi = 0; vi < vk.length; vi++) docState.rows.delete(vk[vi]);
      docState.snapshot.variable.clear();
      docState.frames = docState.frames.filter(function(f) { return !f.isVariable; });
      if (docState.frames.length === 0) {
        docState.active = false;
        figma.ui.postMessage({ type: "live-sync-status", active: false });
      }
      stopVarPolling();
      return;
    }

    // Collect entries (can't iterate and modify Map simultaneously)
    var entries = [];
    docState.snapshot.variable.forEach(function(oldJson, varId) {
      entries.push({ varId: varId, oldJson: oldJson });
    });

    var updated = 0;
    var toRemove = [];

    for (var ei = 0; ei < entries.length; ei++) {
      var entry = entries[ei];
      var rowKey = "var:" + entry.varId;

      if (!docState.rows.has(rowKey)) { toRemove.push(entry.varId); continue; }

      // Check the row node still exists on canvas
      var rowNodeId = docState.rows.get(rowKey);
      if (!figma.getNodeById(rowNodeId)) { toRemove.push(entry.varId); continue; }

      var v;
      try { v = figma.variables.getVariableById(entry.varId); } catch(e) { v = null; }
      if (!v) { toRemove.push(entry.varId); continue; }

      var col;
      try { col = figma.variables.getVariableCollectionById(v.variableCollectionId); } catch(e) { col = null; }
      if (!col) continue;

      // Build current data for this single variable
      var valuesByMode = {};
      for (var mi = 0; mi < col.modes.length; mi++) {
        valuesByMode[col.modes[mi].name] = resolveVariableValue(v, col.modes[mi].modeId, col.modes[mi].name);
      }
      var currentData = {
        id: v.id,
        name: v.name,
        resolvedType: v.resolvedType,
        description: v.description || "",
        valuesByMode: valuesByMode,
      };

      var newJson = JSON.stringify(currentData);
      if (entry.oldJson !== newJson) {
        await rebuildRow(rowKey, "variable", currentData);
        docState.snapshot.variable.set(entry.varId, newJson);
        updated++;
      }
    }

    for (var ri = 0; ri < toRemove.length; ri++) {
      docState.snapshot.variable.delete(toRemove[ri]);
      docState.rows.delete("var:" + toRemove[ri]);
    }

    if (updated > 0) {
      figma.ui.postMessage({ type: "live-sync-update", count: updated, timestamp: new Date().toISOString() });
      figma.notify("Doc updated \u2014 " + updated + " variable" + (updated > 1 ? "s" : "") + " refreshed");
    }
  } catch (err) {
    console.error("[VarPoll] Error:", err);
  } finally {
    _varPollBusy = false;
  }
}

async function lf(fam, sty) {
  const k = fam + "|" + sty;
  if (_lf.has(k)) return true;
  try { await figma.loadFontAsync({ family: fam, style: sty }); _lf.add(k); return true; }
  catch (e) { return false; }
}

// ─── Core helpers ───

// Append child → THEN set sizing (FILL/HUG/FIXED). This order is critical in Figma.
function ac(parent, child, h, v, fixedW) {
  parent.appendChild(child);
  if (h) child.layoutSizingHorizontal = h;
  if (v) child.layoutSizingVertical = v;
  if (h === "FIXED" && fixedW) child.resize(fixedW, child.height || 1);
  if (child.type === "TEXT" && h === "FILL") child.textAutoResize = "HEIGHT";
  return child;
}

// Create text node (no sizing set — caller uses ac())
async function dt(str, sz, sty, col, fam) {
  fam = fam || "Inter"; sty = sty || "Regular";
  if (!(await lf(fam, sty))) { await lf("Inter", "Regular"); fam = "Inter"; sty = "Regular"; }
  const t = figma.createText();
  t.fontName = { family: fam, style: sty };
  t.fontSize = sz;
  t.fills = [{ type: "SOLID", color: col }];
  t.characters = str || " ";
  return t;
}

// Create auto-layout frame (HUG/HUG by default — caller uses ac() to set sizing)
function df(name, dir, gap) {
  const f = figma.createFrame();
  f.name = name || "Frame";
  f.layoutMode = dir || "VERTICAL";
  f.itemSpacing = gap != null ? gap : 0;
  f.fills = [];
  return f;
}

// Badge: HUG/HUG container with padded text
async function dBadge(label, bg, fg) {
  var b = df("Badge", "HORIZONTAL", 0);
  b.paddingLeft = 10; b.paddingRight = 10; b.paddingTop = 5; b.paddingBottom = 5;
  b.cornerRadius = 4;
  b.fills = [{ type: "SOLID", color: bg }];
  ac(b, await dt(label, DC.MIN, "Medium", fg));
  // Force HUG on both axes so parent can't make it FIXED
  b.primaryAxisSizingMode = "AUTO";
  b.counterAxisSizingMode = "AUTO";
  return b;
}

// Format a resolved variable value (from resolveVariableValue/resolveAliasChain)
// into a short display string. Follows alias wrappers down to the concrete value.
function formatRawValue(rv) {
  if (!rv) return null;
  if (rv.type === "alias") return formatRawValue(rv.resolvedValue);
  if (rv.type === "number") return String(Math.round(rv.value * 1000) / 1000);
  if (rv.type === "string") return rv.value;
  if (rv.type === "boolean") return rv.value ? "true" : "false";
  if (rv.type === "color") return (rv.hex || "") + (rv.a != null && rv.a < 1 ? " " + Math.round(rv.a * 100) + "%" : "");
  return null;
}

// Small discreet chip for raw values (pink tint, WCAG AA at 16px)
async function dRawChip(label) {
  var chip = await dBadge(label, DC.rawChipBg, DC.rawValue);
  chip.paddingLeft = 8; chip.paddingRight = 8; chip.paddingTop = 3; chip.paddingBottom = 3;
  return chip;
}

// Style description callout — light blue block with a left accent bar so the
// description stands out from the token/badge noise around it (12.8:1, AAA)
async function dDescCallout(text) {
  var d = df("Description Note", "HORIZONTAL", 0);
  d.paddingLeft = 14; d.paddingRight = 14; d.paddingTop = 10; d.paddingBottom = 10;
  d.cornerRadius = 6;
  d.fills = [{ type: "SOLID", color: DC.descBg }];
  d.strokes = [{ type: "SOLID", color: DC.headerBg }];
  d.strokeLeftWeight = 3; d.strokeTopWeight = 0; d.strokeRightWeight = 0; d.strokeBottomWeight = 0;
  ac(d, await dt(text, DC.MIN, "Regular", DC.text), "FILL", "HUG");
  return d;
}

// ─── Divider ───
function dDiv(parent) {
  const w = df("Divider", "VERTICAL", 0);
  w.paddingLeft = DC.P; w.paddingRight = DC.P;
  const r = figma.createRectangle();
  r.name = "Line"; r.resize(100, 1);
  r.fills = [{ type: "SOLID", color: DC.divider }];
  w.appendChild(r);
  r.layoutSizingHorizontal = "FILL";
  ac(parent, w, "FILL", "HUG");
}

// ─── Header ───
async function dHeader(parent, group) {
  var h = df("Header", "VERTICAL", 12);
  h.fills = [{ type: "SOLID", color: DC.headerBg }];
  h.paddingLeft = DC.P; h.paddingRight = DC.P;
  h.paddingTop = DC.P; h.paddingBottom = DC.P;

  // Title
  ac(h, await dt(group.groupName, 42, "Bold", DC.white), "FILL", "HUG");

  // Counts line: "24 TEXT · 12 COLOUR · 3 EFFECT"
  var counts = [];
  if (group.styles.textStyles.length) counts.push(group.styles.textStyles.length + " text");
  if (group.styles.paintStyles.length) counts.push(group.styles.paintStyles.length + " colour");
  if (group.styles.effectStyles.length) counts.push(group.styles.effectStyles.length + " effect");
  if (group.styles.gridStyles.length) counts.push(group.styles.gridStyles.length + " grid");
  var countLine = await dt(counts.join(" \u00B7 ").toUpperCase(), DC.MIN, "Medium", DC.white);
  countLine.opacity = 0.5;
  countLine.letterSpacing = { value: 2, unit: "PIXELS" };
  ac(h, countLine, "FILL", "HUG");

  // Dynamic description: collect sub-categories from all style names
  var subs = {};
  var allStyles = [].concat(
    group.styles.textStyles || [],
    group.styles.paintStyles || [],
    group.styles.effectStyles || [],
    group.styles.gridStyles || []
  );
  for (var si = 0; si < allStyles.length; si++) {
    var parts = allStyles[si].name.split("/");
    if (parts.length > 1) {
      subs[parts[1].trim()] = true;
    }
  }
  var subList = Object.keys(subs);
  if (subList.length > 0) {
    var descText = "Includes: " + subList.join(", ");
    var descLine = await dt(descText, DC.MIN, "Regular", DC.white);
    descLine.opacity = 0.4;
    ac(h, descLine, "FILL", "HUG");
  }

  // Decorative bar chart
  var deco = df("Deco", "HORIZONTAL", 4);
  deco.counterAxisAlignItems = "MAX";
  var barHeights = [18, 32, 24, 40, 28, 36, 48];
  for (var bi = 0; bi < barHeights.length; bi++) {
    var bar = figma.createRectangle();
    bar.name = "Bar"; bar.resize(8, barHeights[bi]); bar.cornerRadius = 2;
    bar.fills = [{ type: "SOLID", color: DC.white }]; bar.opacity = 0.15;
    deco.appendChild(bar);
  }
  h.appendChild(deco);
  deco.layoutPositioning = "ABSOLUTE";
  deco.x = DC.W - DC.P - 80; deco.y = DC.P;

  ac(parent, h, "FILL", "HUG");
  console.log("[dHeader] w=" + h.width + " h=" + h.height + " hSizing=" + h.layoutSizingHorizontal);
}

// ─── Section title ───
async function dSectionTitle(parent, title) {
  const w = df("Section Title", "HORIZONTAL", 0);
  w.paddingLeft = DC.P; w.paddingRight = DC.P;
  w.paddingTop = 32; w.paddingBottom = 16;
  ac(w, await dt(title, 24, "Bold", DC.text), "FILL", "HUG");
  ac(parent, w, "FILL", "HUG");
}

// ─── Column headers ───
async function dColHeaders(parent) {
  var row = df("Column Headers", "HORIZONTAL", DC.GAP);
  row.paddingLeft = DC.P; row.paddingRight = DC.P; row.paddingBottom = 16;

  // Example header — FIXED width
  var ex = df("Ex Header", "VERTICAL", 0);
  ac(ex, await dt("Example", DC.MIN, "Medium", DC.textSec), "FILL", "HUG");
  ac(row, ex, "FIXED", "HUG", DC.EX);

  // Description header — FILL remaining
  var desc = df("Desc Header", "VERTICAL", 0);
  ac(desc, await dt("Description", DC.MIN, "Medium", DC.textSec), "FILL", "HUG");
  ac(row, desc, "FILL", "HUG");

  // Token header — HUG
  var tok = df("Tok Header", "VERTICAL", 0);
  tok.counterAxisAlignItems = "MAX";
  ac(tok, await dt("Token name", DC.MIN, "Medium", DC.textSec));
  ac(row, tok, "HUG", "HUG");

  ac(parent, row, "FILL", "HUG");
}

// ─── Text row ───
async function dTextRow(parent, style) {
  var row = df("Row \u2014 " + style.name, "HORIZONTAL", DC.GAP);
  row.counterAxisAlignItems = "CENTER";
  row.paddingLeft = DC.P; row.paddingRight = DC.P;
  row.paddingTop = 24; row.paddingBottom = 24;

  // Example column — FIXED 240, HUG height
  var exCol = df("Example", "VERTICAL", 0);
  var exT = await dt("String", style.fontSize, style.fontStyle, DC.text, style.fontFamily);
  if (style.lineHeight && style.lineHeight.unit !== "AUTO") {
    exT.lineHeight = style.lineHeight.unit === "PERCENT"
      ? { value: style.lineHeight.value, unit: "PERCENT" }
      : { value: style.lineHeight.value, unit: "PIXELS" };
  }
  if (style.letterSpacing && style.letterSpacing.value !== 0) {
    exT.letterSpacing = { value: style.letterSpacing.value, unit: style.letterSpacing.unit };
  }
  ac(exCol, exT, "FILL", "HUG");
  ac(row, exCol, "FIXED", "HUG", DC.EX);

  // Description column — FILL width, HUG height
  var descCol = df("Description", "VERTICAL", 14);
  var nameParts = style.name.split("/");
  var shortName = nameParts.length > 1 ? nameParts.slice(1).join(" / ") : nameParts[0];
  ac(descCol, await dt(shortName, DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");
  if (style.description) {
    ac(descCol, await dDescCallout(style.description), "FILL", "HUG");
  }

  // Property badges — FILL width, WRAP
  var badges = df("Properties", "HORIZONTAL", 6);
  badges.layoutWrap = "WRAP"; badges.counterAxisSpacing = 6;
  var addB = async function(label) { ac(badges, await dBadge(label, DC.badgeBg, DC.badgeTxt)); };
  var bv = style.boundVariables || {};
  await addB(bv.fontSize ? "$" + bv.fontSize : "$font-size-" + Math.round(style.fontSize));
  if (style.lineHeight && style.lineHeight.unit !== "AUTO") {
    await addB(bv.lineHeight ? "$" + bv.lineHeight : "$line-height-" + Math.round(style.lineHeight.value));
  }
  await addB(bv.fontFamily ? "$" + bv.fontFamily : "$font-family-" + style.fontFamily.toLowerCase().replace(/\s+/g, "-"));
  var weightBv = bv.fontStyle || bv.fontWeight; // weight can be bound under either key
  await addB(weightBv ? "$" + weightBv : "$font-weight-" + style.fontStyle.toLowerCase().replace(/\s+/g, "-"));
  if (style.letterSpacing) {
    var ls = style.letterSpacing.value;
    await addB(bv.letterSpacing ? "$" + bv.letterSpacing : "$letter-spacing-" + (ls === 0 ? "00" : ls.toFixed(1)));
  }
  ac(descCol, badges, "FILL", "HUG");

  // Responsive modes table — shows EVERY text prop, with per-mode value when bound, or literal otherwise
  if (style.modeNames && style.modeNames.length > 0) {
    var rmModes = style.modeNames;
    var bvModes = style.boundVarModes || {};
    var bvNames = style.boundVariables || {};
    var bvValues = style.boundVarValues || {};

    // Build the canonical list of text-style rows
    var rmRows = [];
    rmRows.push({ label: "font-size", prop: "fontSize", literal: String(Math.round(style.fontSize * 1000) / 1000) });
    if (style.lineHeight && style.lineHeight.unit !== "AUTO") {
      var lhUnit = style.lineHeight.unit === "PERCENT" ? "%" : "px";
      rmRows.push({ label: "line-height", prop: "lineHeight", literal: String(Math.round(style.lineHeight.value * 1000) / 1000) + lhUnit });
    }
    rmRows.push({ label: "font-family", prop: "fontFamily", literal: style.fontFamily });
    // Weight can be bound under fontStyle (string, "Semi Bold") OR fontWeight (number, 600)
    rmRows.push({ label: "font-weight", prop: "fontStyle", altProp: "fontWeight", literal: style.fontStyle });
    if (style.letterSpacing) {
      var lsUnit = style.letterSpacing.unit === "PERCENT" ? "%" : "px";
      var lsv = style.letterSpacing.value;
      rmRows.push({ label: "letter-spacing", prop: "letterSpacing", literal: (lsv === 0 ? "0" : String(Math.round(lsv * 1000) / 1000)) + lsUnit });
    }
    if (style.paragraphSpacing && style.paragraphSpacing > 0) {
      rmRows.push({ label: "paragraph-spacing", prop: "paragraphSpacing", literal: String(style.paragraphSpacing) + "px" });
    }

    // Only render the table when at least one prop is multi-mode bound
    var hasMultiMode = false;
    for (var rmci = 0; rmci < rmRows.length; rmci++) {
      if (bvModes[rmRows[rmci].prop] || (rmRows[rmci].altProp && bvModes[rmRows[rmci].altProp])) { hasMultiMode = true; break; }
    }

    if (hasMultiMode) {
      ac(descCol, await dt("Responsive values", DC.MIN, "Medium", DC.textSec), "FILL", "HUG");

      var rmTable = df("Responsive Table", "VERTICAL", 0);
      rmTable.cornerRadius = 6;
      rmTable.strokes = [{ type: "SOLID", color: DC.divider }];
      rmTable.strokeWeight = 1;
      rmTable.clipsContent = true;

      // Header row
      var rmHdr = df("Header", "HORIZONTAL", 0);
      rmHdr.fills = [{ type: "SOLID", color: DC.badgeBg }];
      var rmH0 = df("Cell", "VERTICAL", 0);
      rmH0.paddingLeft = 14; rmH0.paddingRight = 14; rmH0.paddingTop = 10; rmH0.paddingBottom = 10;
      ac(rmH0, await dt("Property", DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");
      ac(rmHdr, rmH0, "FIXED", "FILL", 160);
      for (var rmh = 0; rmh < rmModes.length; rmh++) {
        var rmHc = df("Cell", "VERTICAL", 0);
        rmHc.paddingLeft = 14; rmHc.paddingRight = 14; rmHc.paddingTop = 10; rmHc.paddingBottom = 10;
        ac(rmHc, await dt(rmModes[rmh], DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");
        ac(rmHdr, rmHc, "FILL", "FILL");
      }
      ac(rmTable, rmHdr, "FILL", "HUG");

      // Data rows — one per text prop
      for (var rmri = 0; rmri < rmRows.length; rmri++) {
        var rmRow = rmRows[rmri];
        // Use the alternative binding key when the primary one has no data
        var rmProp = rmRow.prop;
        if (rmRow.altProp && !bvModes[rmProp] && !bvNames[rmProp] && (bvModes[rmRow.altProp] || bvNames[rmRow.altProp])) {
          rmProp = rmRow.altProp;
        }
        var rmModeMap = bvModes[rmProp];
        var rmFallback = bvNames[rmProp] ? bvNames[rmProp] : rmRow.literal;

        var rmDivLine = figma.createRectangle();
        rmDivLine.name = "Divider"; rmDivLine.resize(100, 1);
        rmDivLine.fills = [{ type: "SOLID", color: DC.divider }];
        ac(rmTable, rmDivLine, "FILL", "FIXED");

        var rmDataRow = df("Row " + rmRow.label, "HORIZONTAL", 0);
        rmDataRow.fills = [{ type: "SOLID", color: DC.white }];

        var rmL = df("Cell", "VERTICAL", 0);
        rmL.paddingLeft = 14; rmL.paddingRight = 14; rmL.paddingTop = 10; rmL.paddingBottom = 10;
        ac(rmL, await dt(rmRow.label, DC.MIN, "Medium", DC.text), "FILL", "HUG");
        ac(rmDataRow, rmL, "FIXED", "FILL", 160);

        for (var rmc = 0; rmc < rmModes.length; rmc++) {
          var rmCell = df("Cell", "VERTICAL", 6);
          rmCell.paddingLeft = 14; rmCell.paddingRight = 14; rmCell.paddingTop = 10; rmCell.paddingBottom = 10;
          var rmTxt = rmFallback;
          var rmRaw = null; // resolved raw value, shown as a chip under the token name
          if (rmModeMap) {
            var rmV = rmModeMap[rmModes[rmc]];
            if (rmV) {
              if (rmV.type === "alias") {
                rmTxt = rmV.aliasName || rmFallback;
                rmRaw = formatRawValue(rmV.resolvedValue);
              }
              else if (rmV.type === "color") rmTxt = (rmV.hex || "#000") + (rmV.a != null && rmV.a < 1 ? " " + Math.round(rmV.a * 100) + "%" : "");
              else if (rmV.type === "number") rmTxt = String(Math.round(rmV.value * 1000) / 1000);
              else if (rmV.type === "string") rmTxt = rmV.value;
              else if (rmV.type === "boolean") rmTxt = rmV.value ? "true" : "false";
            }
          } else if (bvNames[rmProp] && bvValues[rmProp]) {
            // Prop bound to a single-mode collection (e.g. a primitive) — the cell
            // shows the token name, so add its raw value (same across breakpoints)
            rmRaw = formatRawValue(bvValues[rmProp]);
          }
          ac(rmCell, await dt(rmTxt, DC.MIN, "Regular", DC.text), "FILL", "HUG");
          if (rmRaw != null && rmRaw !== "" && rmRaw !== rmTxt) {
            ac(rmCell, await dRawChip(rmRaw));
          }
          ac(rmDataRow, rmCell, "FILL", "FILL");
        }
        ac(rmTable, rmDataRow, "FILL", "HUG");
      }

      ac(descCol, rmTable, "FILL", "HUG");
    }
  }

  ac(row, descCol, "FILL", "HUG");

  // Token column — HUG width, HUG height
  var tokCol = df("Token", "VERTICAL", 0);
  tokCol.counterAxisAlignItems = "MAX";
  var tn = "$text-" + style.name.replace(/\//g, "-").replace(/\s+/g, "-").toLowerCase();
  var tb = await dBadge(tn, DC.headerBg, DC.white);
  tb.cornerRadius = 6;
  ac(tokCol, tb);
  ac(row, tokCol, "HUG", "HUG");

  ac(parent, row, "FILL", "HUG");
  console.log("[dTextRow] " + style.name + " row w=" + row.width + " h=" + row.height + " hSizing=" + row.layoutSizingHorizontal);
  return row;
}

async function dTextSection(parent, styles) {
  await dSectionTitle(parent, "Text Styles");
  await dColHeaders(parent);
  dDiv(parent);
  for (var s = 0; s < styles.length; s++) {
    var row = await dTextRow(parent, styles[s]);
    if (row) {
      docState.rows.set("text:" + styles[s].name, row.id);
      docState.snapshot.text.set(styles[s].name, JSON.stringify(styles[s]));
    }
    dDiv(parent);
  }
}

// ─── Paint row ───
async function dPaintRow(parent, style) {
  var row = df("Row \u2014 " + style.name, "HORIZONTAL", DC.GAP);
  row.counterAxisAlignItems = "CENTER";
  row.paddingLeft = DC.P; row.paddingRight = DC.P;
  row.paddingTop = 20; row.paddingBottom = 20;

  var paint = style.paints[0];
  var sw = figma.createRectangle();
  sw.name = "Swatch"; sw.resize(56, 56); sw.cornerRadius = 8;

  if (paint && paint.type === "SOLID" && paint.color) {
    // Solid fill
    sw.fills = [{ type: "SOLID", color: { r: paint.color.r / 255, g: paint.color.g / 255, b: paint.color.b / 255 }, opacity: paint.color.a }];
  } else if (paint && paint.type.startsWith("GRADIENT_") && paint.gradientStops) {
    // Gradient fill — rebuild stops for Figma API (0-255 → 0-1)
    var stops = [];
    for (var gi = 0; gi < paint.gradientStops.length; gi++) {
      var gs = paint.gradientStops[gi];
      stops.push({
        position: gs.position,
        color: { r: gs.color.r / 255, g: gs.color.g / 255, b: gs.color.b / 255, a: gs.color.a != null ? gs.color.a : 1 },
      });
    }
    var gradFill = { type: paint.type, gradientStops: stops };
    if (paint.gradientTransform) {
      gradFill.gradientTransform = paint.gradientTransform;
    } else {
      gradFill.gradientTransform = [[1, 0, 0], [0, 1, 0]];
    }
    sw.fills = [gradFill];
  } else {
    sw.fills = [{ type: "SOLID", color: { r: 0.85, g: 0.85, b: 0.85 } }];
  }
  // Subtle border so swatch is visible even on white
  sw.strokes = [{ type: "SOLID", color: DC.divider }]; sw.strokeWeight = 1;
  row.appendChild(sw);

  // Description — FILL width
  var descCol = df("Description", "VERTICAL", 8);
  var nameParts = style.name.split("/");
  var shortName = nameParts.length > 1 ? nameParts.slice(1).join(" / ") : nameParts[0];
  ac(descCol, await dt(shortName, DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");

  if (paint && paint.type === "SOLID" && paint.color) {
    var alphaStr = paint.color.a < 1 ? " \u00B7 " + Math.round(paint.color.a * 100) + "%" : "";
    ac(descCol, await dt(paint.color.hex + alphaStr, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
  } else if (paint && paint.type.startsWith("GRADIENT_") && paint.gradientStops) {
    // Gradient type label
    var gType = paint.type.replace("GRADIENT_", "").toLowerCase();
    gType = gType.charAt(0).toUpperCase() + gType.slice(1);
    ac(descCol, await dt(gType + " gradient", DC.MIN, "Regular", DC.textSec), "FILL", "HUG");

    // Colour stops as badges with dots
    var stopsRow = df("Stops", "HORIZONTAL", 6);
    stopsRow.layoutWrap = "WRAP"; stopsRow.counterAxisSpacing = 6;
    for (var si = 0; si < paint.gradientStops.length; si++) {
      var stop = paint.gradientStops[si];
      ac(stopsRow, await dColorBadge({
        hex: stop.color.hex,
        r: stop.color.r,
        g: stop.color.g,
        b: stop.color.b,
        a: stop.color.a
      }));
      ac(stopsRow, await dBadge(Math.round(stop.position * 100) + "%", DC.badgeBg, DC.badgeTxt));
    }
    ac(descCol, stopsRow, "FILL", "HUG");
  }

  if (style.description) {
    ac(descCol, await dDescCallout(style.description), "FILL", "HUG");
  }
  ac(row, descCol, "FILL", "HUG");

  // Token — HUG
  var tokCol = df("Token", "VERTICAL", 0);
  tokCol.counterAxisAlignItems = "MAX";
  var tn = "$color-" + style.name.replace(/\//g, "-").replace(/\s+/g, "-").toLowerCase();
  var tb = await dBadge(tn, DC.headerBg, DC.white);
  tb.cornerRadius = 6;
  ac(tokCol, tb);
  ac(row, tokCol, "HUG", "HUG");

  ac(parent, row, "FILL", "HUG");
  return row;
}

async function dPaintSection(parent, styles) {
  await dSectionTitle(parent, "Colour Styles");
  dDiv(parent);
  for (var s = 0; s < styles.length; s++) {
    var row = await dPaintRow(parent, styles[s]);
    if (row) {
      docState.rows.set("paint:" + styles[s].name, row.id);
      docState.snapshot.paint.set(styles[s].name, JSON.stringify(styles[s]));
    }
    dDiv(parent);
  }
}

// ─── Effect: colour swatch badge (small rect + hex text) ───
async function dColorBadge(color) {
  var wrap = df("Colour", "HORIZONTAL", 6);
  wrap.paddingLeft = 8; wrap.paddingRight = 10; wrap.paddingTop = 5; wrap.paddingBottom = 5;
  wrap.cornerRadius = 4;
  wrap.fills = [{ type: "SOLID", color: DC.badgeBg }];
  wrap.counterAxisAlignItems = "CENTER";
  // Small colour dot
  var dot = figma.createRectangle();
  dot.name = "Dot"; dot.resize(14, 14); dot.cornerRadius = 3;
  dot.fills = [{ type: "SOLID", color: { r: color.r / 255, g: color.g / 255, b: color.b / 255 }, opacity: color.a != null ? color.a : 1 }];
  dot.strokes = [{ type: "SOLID", color: DC.divider }]; dot.strokeWeight = 1;
  wrap.appendChild(dot);
  var label = color.hex || "#000000";
  if (color.a != null && color.a < 1) label += " " + Math.round(color.a * 100) + "%";
  ac(wrap, await dt(label, DC.MIN, "Medium", DC.badgeTxt));
  // Force HUG on both axes
  wrap.primaryAxisSizingMode = "AUTO";
  wrap.counterAxisSizingMode = "AUTO";
  return wrap;
}

// ─── Effect row ───
async function dEffectRow(parent, style) {
  var row = df("Row \u2014 " + style.name, "HORIZONTAL", DC.GAP);
  row.counterAxisAlignItems = "CENTER";
  row.paddingLeft = DC.P; row.paddingRight = DC.P;
  row.paddingTop = 20; row.paddingBottom = 20;

  // Preview rectangle with actual effects applied
  var pv = figma.createRectangle();
  pv.name = "Preview"; pv.resize(56, 56); pv.cornerRadius = 8;
  pv.fills = [{ type: "SOLID", color: DC.white }];
  var fx = [];
  for (var ei = 0; ei < style.effects.length; ei++) {
    var e = style.effects[ei];
    if ((e.type === "DROP_SHADOW" || e.type === "INNER_SHADOW") && e.visible) {
      fx.push({
        type: e.type, visible: true,
        color: e.color ? { r: e.color.r / 255, g: e.color.g / 255, b: e.color.b / 255, a: e.color.a != null ? e.color.a : 0.25 } : { r: 0, g: 0, b: 0, a: 0.25 },
        offset: { x: e.offset ? e.offset.x : 0, y: e.offset ? e.offset.y : 0 },
        radius: e.radius || 0, spread: e.spread || 0, blendMode: e.blendMode || "NORMAL",
      });
    }
    if ((e.type === "LAYER_BLUR" || e.type === "BACKGROUND_BLUR") && e.visible) {
      fx.push({ type: e.type, visible: true, radius: e.radius || 0 });
    }
  }
  if (fx.length) pv.effects = fx;
  row.appendChild(pv);

  // Description column — FILL
  var descCol = df("Description", "VERTICAL", 10);
  var nameParts = style.name.split("/");
  var shortName = nameParts.length > 1 ? nameParts.slice(1).join(" / ") : nameParts[0];
  ac(descCol, await dt(shortName, DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");
  if (style.description) {
    ac(descCol, await dDescCallout(style.description), "FILL", "HUG");
  }

  // One detail block per effect
  for (var ei2 = 0; ei2 < style.effects.length; ei2++) {
    var eff = style.effects[ei2];
    if (!eff.visible) continue;

    var effectBlock = df("Effect " + (ei2 + 1), "VERTICAL", 6);
    var typeName = eff.type.toLowerCase().replace(/_/g, " ");
    typeName = typeName.charAt(0).toUpperCase() + typeName.slice(1);
    ac(effectBlock, await dt(typeName, DC.MIN, "Medium", DC.text), "FILL", "HUG");

    var props = df("Props", "HORIZONTAL", 6);
    props.layoutWrap = "WRAP"; props.counterAxisSpacing = 6;

    if (eff.type === "DROP_SHADOW" || eff.type === "INNER_SHADOW") {
      // Colour badge with swatch
      if (eff.color) {
        ac(props, await dColorBadge(eff.color));
      }
      // Position
      var ox = eff.offset ? eff.offset.x : 0;
      var oy = eff.offset ? eff.offset.y : 0;
      ac(props, await dBadge("X: " + ox + "  Y: " + oy, DC.badgeBg, DC.badgeTxt));
      // Blur
      ac(props, await dBadge("Blur: " + (eff.radius != null ? eff.radius : 0), DC.badgeBg, DC.badgeTxt));
      // Spread
      ac(props, await dBadge("Spread: " + (eff.spread != null ? eff.spread : 0), DC.badgeBg, DC.badgeTxt));
      // Blend mode (only if not Normal)
      if (eff.blendMode && eff.blendMode !== "NORMAL" && eff.blendMode !== "PASS_THROUGH") {
        ac(props, await dBadge("Blend: " + eff.blendMode.toLowerCase().replace(/_/g, " "), DC.badgeBg, DC.badgeTxt));
      }
    }
    if (eff.type === "LAYER_BLUR" || eff.type === "BACKGROUND_BLUR") {
      ac(props, await dBadge("Radius: " + (eff.radius || 0), DC.badgeBg, DC.badgeTxt));
    }

    ac(effectBlock, props, "FILL", "HUG");
    ac(descCol, effectBlock, "FILL", "HUG");
  }

  ac(row, descCol, "FILL", "HUG");

  // Token — HUG
  var tokCol = df("Token", "VERTICAL", 0);
  tokCol.counterAxisAlignItems = "MAX";
  var tn = "$effect-" + style.name.replace(/\//g, "-").replace(/\s+/g, "-").toLowerCase();
  var tb = await dBadge(tn, DC.headerBg, DC.white);
  tb.cornerRadius = 6;
  ac(tokCol, tb);
  ac(row, tokCol, "HUG", "HUG");

  ac(parent, row, "FILL", "HUG");
  return row;
}

// ─── Footer ───
async function dFooter(parent) {
  var now = new Date();
  var dd = String(now.getDate()).padStart(2, "0");
  var mm = String(now.getMonth() + 1).padStart(2, "0");
  var yyyy = now.getFullYear();
  var hh = String(now.getHours()).padStart(2, "0");
  var min = String(now.getMinutes()).padStart(2, "0");
  var timestamp = dd + "/" + mm + "/" + yyyy + " at " + hh + ":" + min;

  var foot = df("Footer", "HORIZONTAL", 0);
  foot.paddingLeft = DC.P; foot.paddingRight = DC.P;
  foot.paddingTop = 32; foot.paddingBottom = 32;
  foot.counterAxisAlignItems = "CENTER";

  ac(foot, await dt("Generated on " + timestamp + "  \u00B7  DS Styles Extractor", DC.MIN, "Regular", DC.textSec), "FILL", "HUG");

  ac(parent, foot, "FILL", "HUG");
}

async function dEffectSection(parent, styles) {
  await dSectionTitle(parent, "Effect Styles");
  dDiv(parent);
  for (var s = 0; s < styles.length; s++) {
    var row = await dEffectRow(parent, styles[s]);
    if (row) {
      docState.rows.set("effect:" + styles[s].name, row.id);
      docState.snapshot.effect.set(styles[s].name, JSON.stringify(styles[s]));
    }
    dDiv(parent);
  }
}

// ─── Variable Documentation ───

// Variable header
async function dVariableHeader(parent, collectionName, modes, variableCount) {
  var h = df("Header", "VERTICAL", 12);
  h.fills = [{ type: "SOLID", color: DC.headerBg }];
  h.paddingLeft = DC.P; h.paddingRight = DC.P;
  h.paddingTop = DC.P; h.paddingBottom = DC.P;

  ac(h, await dt(collectionName + " Variables", 42, "Bold", DC.white), "FILL", "HUG");

  var countLine = await dt(
    variableCount + " VARIABLES \u00B7 " + modes.length + " MODE" + (modes.length > 1 ? "S" : ""),
    DC.MIN, "Medium", DC.white
  );
  countLine.opacity = 0.5;
  countLine.letterSpacing = { value: 2, unit: "PIXELS" };
  ac(h, countLine, "FILL", "HUG");

  var modeNames = [];
  for (var mi = 0; mi < modes.length; mi++) modeNames.push(modes[mi].name);
  var descLine = await dt("Modes: " + modeNames.join(", "), DC.MIN, "Regular", DC.white);
  descLine.opacity = 0.4;
  ac(h, descLine, "FILL", "HUG");

  // Decorative bars
  var deco = df("Deco", "HORIZONTAL", 4);
  deco.counterAxisAlignItems = "MAX";
  var barHeights = [18, 32, 24, 40, 28, 36, 48];
  for (var bi = 0; bi < barHeights.length; bi++) {
    var bar = figma.createRectangle();
    bar.name = "Bar"; bar.resize(8, barHeights[bi]); bar.cornerRadius = 2;
    bar.fills = [{ type: "SOLID", color: DC.white }]; bar.opacity = 0.15;
    deco.appendChild(bar);
  }
  h.appendChild(deco);
  deco.layoutPositioning = "ABSOLUTE";
  deco.x = DC.W - DC.P - 80; deco.y = DC.P;

  ac(parent, h, "FILL", "HUG");
}

// Variable column headers (no preview column)
async function dVarColHeaders(parent) {
  var row = df("Column Headers", "HORIZONTAL", DC.GAP);
  row.paddingLeft = DC.P; row.paddingRight = DC.P; row.paddingBottom = 16;

  var desc = df("Desc Header", "VERTICAL", 0);
  ac(desc, await dt("Variable", DC.MIN, "Medium", DC.textSec), "FILL", "HUG");
  ac(row, desc, "FILL", "HUG");

  var tok = df("Tok Header", "VERTICAL", 0);
  tok.counterAxisAlignItems = "MAX";
  ac(tok, await dt("Token path", DC.MIN, "Medium", DC.textSec));
  ac(row, tok, "HUG", "HUG");

  ac(parent, row, "FILL", "HUG");
}

// Variable row
async function dVariableRow(parent, variable, modes) {
  var row = df("Row \u2014 " + variable.name, "HORIZONTAL", DC.GAP);
  row.counterAxisAlignItems = "CENTER";
  row.paddingLeft = DC.P; row.paddingRight = DC.P;
  row.paddingTop = 20; row.paddingBottom = 20;

  // Description column
  var descCol = df("Description", "VERTICAL", 6);
  var nameParts = variable.name.split("/");
  var shortName = nameParts.length > 1 ? nameParts.slice(1).join(" / ") : nameParts[0];
  ac(descCol, await dt(shortName, DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");

  if (variable.description) {
    ac(descCol, await dt(variable.description, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
  }

  // Type + mode value badges
  var badges = df("Values", "HORIZONTAL", 6);
  badges.layoutWrap = "WRAP"; badges.counterAxisSpacing = 6;
  ac(badges, await dBadge(variable.resolvedType.toLowerCase(), DC.badgeBg, DC.badgeTxt));

  for (var mi2 = 0; mi2 < modes.length; mi2++) {
    var modeVal = variable.valuesByMode[modes[mi2].name];
    if (!modeVal) continue;

    if (modeVal.type === "color") {
      // Mode-labelled colour badge
      var cWrap = df("Mode", "HORIZONTAL", 6);
      cWrap.paddingLeft = 8; cWrap.paddingRight = 10; cWrap.paddingTop = 5; cWrap.paddingBottom = 5;
      cWrap.cornerRadius = 4;
      cWrap.fills = [{ type: "SOLID", color: DC.badgeBg }];
      cWrap.counterAxisAlignItems = "CENTER";
      ac(cWrap, await dt(modes[mi2].name + ":", DC.MIN, "Medium", DC.textSec));
      var cdot = figma.createRectangle();
      cdot.name = "Dot"; cdot.resize(14, 14); cdot.cornerRadius = 3;
      cdot.fills = [{ type: "SOLID", color: { r: modeVal.r / 255, g: modeVal.g / 255, b: modeVal.b / 255 }, opacity: modeVal.a != null ? modeVal.a : 1 }];
      cdot.strokes = [{ type: "SOLID", color: DC.divider }]; cdot.strokeWeight = 1;
      cWrap.appendChild(cdot);
      ac(cWrap, await dt(modeVal.hex, DC.MIN, "Medium", DC.badgeTxt));
      cWrap.primaryAxisSizingMode = "AUTO";
      cWrap.counterAxisSizingMode = "AUTO";
      ac(badges, cWrap);
    } else if (modeVal.type === "alias") {
      ac(badges, await dBadge(modes[mi2].name + ": \u2192 " + modeVal.aliasName, DC.badgeBg, DC.badgeTxt));
    } else {
      var valStr = "";
      if (modeVal.type === "number") valStr = String(modeVal.value);
      else if (modeVal.type === "string") valStr = "\"" + modeVal.value + "\"";
      else if (modeVal.type === "boolean") valStr = String(modeVal.value);
      ac(badges, await dBadge(modes[mi2].name + ": " + valStr, DC.badgeBg, DC.badgeTxt));
    }
  }

  ac(descCol, badges, "FILL", "HUG");
  ac(row, descCol, "FILL", "HUG");

  // Token column
  var tokCol = df("Token", "VERTICAL", 0);
  tokCol.counterAxisAlignItems = "MAX";
  var tn = "$" + variable.name.replace(/\//g, "-").replace(/\s+/g, "-").toLowerCase();
  var tb = await dBadge(tn, DC.headerBg, DC.white);
  tb.cornerRadius = 6;
  ac(tokCol, tb);
  ac(row, tokCol, "HUG", "HUG");

  ac(parent, row, "FILL", "HUG");
  return row;
}

// ─── Debug: walk tree and log layout properties ───
function debugTree(node, depth) {
  depth = depth || 0;
  var pad = "";
  for (var d = 0; d < depth; d++) pad += "  ";
  var info = pad + node.name + " (" + node.type + ") ";
  info += "w=" + Math.round(node.width) + " h=" + Math.round(node.height);
  if (node.layoutMode && node.layoutMode !== "NONE") {
    info += " layout=" + node.layoutMode;
    info += " gap=" + node.itemSpacing;
  }
  if (node.layoutSizingHorizontal) info += " hSizing=" + node.layoutSizingHorizontal;
  if (node.layoutSizingVertical) info += " vSizing=" + node.layoutSizingVertical;
  if (node.layoutGrow) info += " grow=" + node.layoutGrow;
  if (node.textAutoResize) info += " textResize=" + node.textAutoResize;
  if (node.layoutPositioning === "ABSOLUTE") info += " ABSOLUTE";
  console.log(info);
  if ("children" in node) {
    for (var c = 0; c < node.children.length; c++) {
      debugTree(node.children[c], depth + 1);
    }
  }
}

async function generateDocumentation(payload) {
  var groups = payload.groups;

  // Reset style-related live sync state (preserve variable state)
  docState.active = false;
  docState.frames = docState.frames.filter(function(f) { return f.isVariable; });
  var _sKeys = [];
  docState.rows.forEach(function(v, k) { if (!k.startsWith("var:")) _sKeys.push(k); });
  for (var _sk = 0; _sk < _sKeys.length; _sk++) docState.rows.delete(_sKeys[_sk]);
  docState.snapshot.paint.clear();
  docState.snapshot.text.clear();
  docState.snapshot.effect.clear();

  figma.ui.postMessage({ type: "doc-progress", status: "starting", total: groups.length });

  _lf.clear();
  await lf("Inter", "Regular");
  await lf("Inter", "Medium");
  await lf("Inter", "Semi Bold");
  await lf("Inter", "Bold");

  var page = figma.currentPage;

  // Place new content to the right of existing content
  var startX = 0;
  for (var n = 0; n < page.children.length; n++) {
    var ch = page.children[n];
    var right = ch.x + ch.width;
    if (right > startX) startX = right;
  }
  startX += 200;

  // Wrapper: HORIZONTAL auto-layout with 111px gap when multiple groups
  var wrapper = null;
  if (groups.length > 1) {
    wrapper = figma.createFrame();
    wrapper.name = "Documentation";
    wrapper.layoutMode = "HORIZONTAL";
    wrapper.primaryAxisSizingMode = "AUTO";
    wrapper.counterAxisSizingMode = "AUTO";
    wrapper.itemSpacing = 111;
    wrapper.fills = [];
    page.appendChild(wrapper);
    wrapper.x = startX;
    wrapper.y = 0;
  }

  var allMains = [];

  for (var i = 0; i < groups.length; i++) {
    var g = groups[i];
    figma.ui.postMessage({ type: "doc-progress", status: "generating", current: i + 1, total: groups.length, name: g.groupName });
    console.log("── Generating doc for: " + g.groupName + " ──");

    // Each group frame: VERTICAL, FIXED width 1400, HUG height
    var main = figma.createFrame();
    main.name = g.groupName + " Documentation";
    main.layoutMode = "VERTICAL";
    main.counterAxisSizingMode = "FIXED";
    main.primaryAxisSizingMode = "AUTO";
    main.resize(DC.W, 100);
    main.fills = [{ type: "SOLID", color: DC.white }];
    main.itemSpacing = 0;

    if (wrapper) {
      wrapper.appendChild(main);
    } else {
      page.appendChild(main);
      main.x = startX;
      main.y = 0;
    }

    await dHeader(main, g);
    if (g.styles.textStyles.length) await dTextSection(main, g.styles.textStyles);
    if (g.styles.paintStyles.length) await dPaintSection(main, g.styles.paintStyles);
    if (g.styles.effectStyles.length) await dEffectSection(main, g.styles.effectStyles);

    dDiv(main);
    await dFooter(main);

    allMains.push(main);
    docState.frames.push({ frameId: main.id, groupName: g.groupName, isVariable: false });

    console.log("── TREE DUMP: " + g.groupName + " ──");
    debugTree(main);
    console.log("── END ──");
  }

  var viewTarget = wrapper ? [wrapper] : allMains;
  if (viewTarget.length) {
    figma.viewport.scrollAndZoomIntoView(viewTarget);
  }

  // Activate live sync
  docState.active = true;
  figma.ui.postMessage({ type: "live-sync-status", active: true });

  figma.ui.postMessage({ type: "doc-progress", status: "done", pageCount: groups.length });
  figma.notify("Documentation generated \u2014 " + groups.length + " group" + (groups.length > 1 ? "s" : ""));
}

// ─── Variable Documentation Generator ───

async function generateVariableDocumentation(payload) {
  var modes = payload.modes;
  var groups = payload.groups; // [{ groupName, variables }]
  var collectionName = payload.collectionName || "Variables";

  // Clear variable-related state only (preserve style state)
  var _vKeys = [];
  docState.rows.forEach(function(v, k) { if (k.startsWith("var:")) _vKeys.push(k); });
  for (var _vk = 0; _vk < _vKeys.length; _vk++) docState.rows.delete(_vKeys[_vk]);
  docState.snapshot.variable.clear();
  docState.frames = docState.frames.filter(function(f) { return !f.isVariable; });
  docState.varModes = modes;

  var totalVars = 0;
  for (var tv = 0; tv < groups.length; tv++) totalVars += groups[tv].variables.length;

  figma.ui.postMessage({ type: "doc-progress", status: "starting", total: groups.length, isVariable: true });

  _lf.clear();
  await lf("Inter", "Regular");
  await lf("Inter", "Medium");
  await lf("Inter", "Semi Bold");
  await lf("Inter", "Bold");

  var page = figma.currentPage;

  // Place to the right of existing content
  var startX = 0;
  for (var n = 0; n < page.children.length; n++) {
    var ch = page.children[n];
    var right = ch.x + ch.width;
    if (right > startX) startX = right;
  }
  startX += 200;

  // Wrapper: HORIZONTAL auto-layout when multiple groups
  var wrapper = null;
  if (groups.length > 1) {
    wrapper = figma.createFrame();
    wrapper.name = collectionName + " Variables";
    wrapper.layoutMode = "HORIZONTAL";
    wrapper.primaryAxisSizingMode = "AUTO";
    wrapper.counterAxisSizingMode = "AUTO";
    wrapper.itemSpacing = 111;
    wrapper.fills = [];
    page.appendChild(wrapper);
    wrapper.x = startX;
    wrapper.y = 0;
  }

  var allMains = [];
  var rowCount = 0;

  for (var gi = 0; gi < groups.length; gi++) {
    var g = groups[gi];
    figma.ui.postMessage({ type: "doc-progress", status: "generating", current: gi + 1, total: groups.length, name: g.groupName, isVariable: true });
    console.log("── Generating variable doc for group: " + g.groupName + " (" + g.variables.length + " vars) ──");

    // Each group gets its own frame
    var main = figma.createFrame();
    main.name = g.groupName + " Variables";
    main.layoutMode = "VERTICAL";
    main.counterAxisSizingMode = "FIXED";
    main.primaryAxisSizingMode = "AUTO";
    main.resize(DC.W, 100);
    main.fills = [{ type: "SOLID", color: DC.white }];
    main.itemSpacing = 0;

    if (wrapper) {
      wrapper.appendChild(main);
    } else {
      page.appendChild(main);
      main.x = startX;
      main.y = 0;
    }

    await dVariableHeader(main, g.groupName, modes, g.variables.length);

    // Sub-group variables by the next path segment after the group depth
    var depth = g.depth || 1;
    var subGroups = {};
    var subOrder = [];
    for (var vi = 0; vi < g.variables.length; vi++) {
      var parts = g.variables[vi].name.split("/");
      var subName = parts.length > depth + 1 ? parts[depth] : "General";
      if (!subGroups[subName]) { subGroups[subName] = []; subOrder.push(subName); }
      subGroups[subName].push(g.variables[vi]);
    }
    subOrder.sort();

    for (var si = 0; si < subOrder.length; si++) {
      var sn = subOrder[si];
      var sVars = subGroups[sn];

      await dSectionTitle(main, sn + " (" + sVars.length + ")");
      await dVarColHeaders(main);
      dDiv(main);

      for (var ri = 0; ri < sVars.length; ri++) {
        var v = sVars[ri];
        var row = await dVariableRow(main, v, modes);
        if (row) {
          docState.rows.set("var:" + v.id, row.id);
          // Store only the fields we compare in performVariablePoll
          var snapData = { id: v.id, name: v.name, resolvedType: v.resolvedType, description: v.description || "", valuesByMode: v.valuesByMode };
          docState.snapshot.variable.set(v.id, JSON.stringify(snapData));
        }
        dDiv(main);
        rowCount++;
      }
    }

    dDiv(main);
    await dFooter(main);

    allMains.push(main);
    docState.frames.push({ frameId: main.id, groupName: g.groupName, isVariable: true });
  }

  var viewTarget = wrapper ? [wrapper] : allMains;
  if (viewTarget.length) {
    figma.viewport.scrollAndZoomIntoView(viewTarget);
  }

  docState.active = true;
  startVarPolling();
  figma.ui.postMessage({ type: "live-sync-status", active: true });

  figma.ui.postMessage({ type: "doc-progress", status: "done", pageCount: groups.length, isVariable: true, varCount: totalVars });
  figma.notify("Variable documentation generated \u2014 " + totalVars + " variables in " + groups.length + " group" + (groups.length > 1 ? "s" : ""));
}

// ─── Live Sync: Rebuild a single row ───

async function rebuildRow(key, type, styleData) {
  var nodeId = docState.rows.get(key);
  var oldNode = figma.getNodeById(nodeId);
  if (!oldNode || !oldNode.parent) {
    docState.rows.delete(key);
    return;
  }

  var parent = oldNode.parent;
  var idx = -1;
  for (var i = 0; i < parent.children.length; i++) {
    if (parent.children[i].id === oldNode.id) { idx = i; break; }
  }

  // Remove old row
  oldNode.remove();

  // Create new row (appended to end of parent by the function)
  var newRow;
  if (type === "paint") newRow = await dPaintRow(parent, styleData);
  else if (type === "text") newRow = await dTextRow(parent, styleData);
  else if (type === "effect") newRow = await dEffectRow(parent, styleData);
  else if (type === "variable") newRow = await dVariableRow(parent, styleData, docState.varModes);

  if (!newRow) return;

  // Move to correct position (where the old row was)
  if (idx >= 0 && idx < parent.children.length - 1) {
    parent.insertChild(idx, newRow);
  }

  // Update stored node ID
  docState.rows.set(key, newRow.id);
}

// ─── Live Sync: Diff and update ───

async function performLiveUpdate() {
  if (_isLiveUpdating || !docState.active) return;
  _isLiveUpdating = true;

  try {
    // Verify at least one doc frame still exists
    var anyAlive = false;
    for (var i = 0; i < docState.frames.length; i++) {
      if (figma.getNodeById(docState.frames[i].frameId)) { anyAlive = true; break; }
    }
    if (!anyAlive) {
      docState.active = false;
      stopVarPolling();
      figma.ui.postMessage({ type: "live-sync-status", active: false });
      return;
    }

    // Re-extract current styles
    var current = extractAllStyles();
    var updated = 0;

    // Check paint styles
    for (var p = 0; p < current.paintStyles.length; p++) {
      var ps = current.paintStyles[p];
      var pKey = "paint:" + ps.name;
      var pOld = docState.snapshot.paint.get(ps.name);
      var pNew = JSON.stringify(ps);
      if (pOld && pOld !== pNew && docState.rows.has(pKey)) {
        await rebuildRow(pKey, "paint", ps);
        docState.snapshot.paint.set(ps.name, pNew);
        updated++;
      }
    }

    // Check text styles
    for (var t = 0; t < current.textStyles.length; t++) {
      var ts = current.textStyles[t];
      var tKey = "text:" + ts.name;
      var tOld = docState.snapshot.text.get(ts.name);
      var tNew = JSON.stringify(ts);
      if (tOld && tOld !== tNew && docState.rows.has(tKey)) {
        await rebuildRow(tKey, "text", ts);
        docState.snapshot.text.set(ts.name, tNew);
        updated++;
      }
    }

    // Check effect styles
    for (var e = 0; e < current.effectStyles.length; e++) {
      var es = current.effectStyles[e];
      var eKey = "effect:" + es.name;
      var eOld = docState.snapshot.effect.get(es.name);
      var eNew = JSON.stringify(es);
      if (eOld && eOld !== eNew && docState.rows.has(eKey)) {
        await rebuildRow(eKey, "effect", es);
        docState.snapshot.effect.set(es.name, eNew);
        updated++;
      }
    }

    // Variables are handled by performVariablePoll (lightweight per-ID check)

    if (updated > 0) {
      figma.ui.postMessage({ type: "live-sync-update", count: updated, timestamp: new Date().toISOString() });
      figma.notify("Doc updated \u2014 " + updated + " item" + (updated > 1 ? "s" : "") + " refreshed");
    }
  } catch (err) {
    console.error("[LiveSync] Error:", err);
  } finally {
    _isLiveUpdating = false;
  }
}

// ─── Document change listener ───

figma.on("documentchange", function (event) {
  if (_isLiveUpdating || !docState.active) return;

  // Only react to style events; variable changes are handled by polling
  var styleChange = false;
  for (var i = 0; i < event.documentChanges.length; i++) {
    var ch = event.documentChanges[i];
    if (ch.type === "STYLE_PROPERTY_CHANGE" ||
        ch.type === "STYLE_CREATE" ||
        ch.type === "STYLE_DELETE") {
      styleChange = true;
      break;
    }
  }

  if (!styleChange) return;

  if (_liveTimer) clearTimeout(_liveTimer);
  _liveTimer = setTimeout(function () { performLiveUpdate(); }, 500);
});

// ─── Detect existing documentation on startup ───

function detectExistingDocs() {
  var page = figma.currentPage;
  var styles = extractAllStyles();
  var variables = extractAllVariables();

  // Build lookup sets by name
  var paintNames = new Set();
  var textNames = new Set();
  var effectNames = new Set();
  var variableNames = new Set();
  for (var i = 0; i < styles.paintStyles.length; i++) paintNames.add(styles.paintStyles[i].name);
  for (var i2 = 0; i2 < styles.textStyles.length; i2++) textNames.add(styles.textStyles[i2].name);
  for (var i3 = 0; i3 < styles.effectStyles.length; i3++) effectNames.add(styles.effectStyles[i3].name);

  // Build variable name lookup + store modes for rebuild
  var varLookup = {};
  if (variables.collections) {
    for (var vc = 0; vc < variables.collections.length; vc++) {
      var col = variables.collections[vc];
      for (var vv = 0; vv < col.variables.length; vv++) {
        variableNames.add(col.variables[vv].name);
        varLookup[col.variables[vv].name] = col.variables[vv];
      }
      // Store modes if this looks like our tracked collection
      if (col.name.toLowerCase() === "components" && docState.varModes.length === 0) {
        docState.varModes = col.modes;
      }
    }
  }

  // Scan top-level children for doc frames
  for (var n = 0; n < page.children.length; n++) {
    scanDocFrame(page.children[n], paintNames, textNames, effectNames, variableNames, varLookup, styles);
  }

  if (docState.frames.length > 0) {
    docState.active = true;
    figma.ui.postMessage({ type: "live-sync-status", active: true });
    console.log("[DetectDocs] Found " + docState.frames.length + " doc frame(s), " + docState.rows.size + " rows mapped");

    // Start variable polling if variable docs were detected
    if (docState.snapshot.variable.size > 0) {
      startVarPolling();
      console.log("[DetectDocs] Variable docs found (" + docState.snapshot.variable.size + " vars), polling started");
    }
  }
}

function scanDocFrame(node, paintNames, textNames, effectNames, variableNames, varLookup, styles) {
  if (node.type !== "FRAME") return;

  // Wrapper frame ("Documentation") — recurse into children
  if (node.name === "Documentation" && "children" in node) {
    for (var w = 0; w < node.children.length; w++) {
      scanDocFrame(node.children[w], paintNames, textNames, effectNames, variableNames, varLookup, styles);
    }
    return;
  }

  // Variable wrapper frame (horizontal layout, name ends with " Variables") — recurse
  if (node.name.endsWith(" Variables") && node.layoutMode === "HORIZONTAL" && "children" in node) {
    for (var w2 = 0; w2 < node.children.length; w2++) {
      scanDocFrame(node.children[w2], paintNames, textNames, effectNames, variableNames, varLookup, styles);
    }
    return;
  }

  // Check if this is a variable doc frame (name ends with " Variables")
  var isVarFrame = node.name.endsWith(" Variables");

  // Check if this is a style doc frame (name ends with " Documentation")
  var isStyleFrame = node.name.endsWith(" Documentation");

  if (!isVarFrame && !isStyleFrame) return;

  if (isStyleFrame) {
    var groupName = node.name.replace(" Documentation", "");
    docState.frames.push({ frameId: node.id, groupName: groupName, isVariable: false });
  } else {
    var colName = node.name.replace(" Variables", "");
    docState.frames.push({ frameId: node.id, groupName: colName, isVariable: true });
  }

  // Walk children looking for rows named "Row — <name>"
  var prefix = "Row \u2014 ";
  for (var c = 0; c < node.children.length; c++) {
    var child = node.children[c];
    if (child.type !== "FRAME" || !child.name.startsWith(prefix)) continue;

    var styleName = child.name.substring(prefix.length);

    if (isVarFrame && variableNames.has(styleName) && varLookup[styleName]) {
      // Variable row — store by variable ID for rename detection
      var vl = varLookup[styleName];
      docState.rows.set("var:" + vl.id, child.id);
      var snapData = { id: vl.id, name: vl.name, resolvedType: vl.resolvedType, description: vl.description || "", valuesByMode: vl.valuesByMode };
      docState.snapshot.variable.set(vl.id, JSON.stringify(snapData));
    } else if (paintNames.has(styleName)) {
      docState.rows.set("paint:" + styleName, child.id);
      for (var p = 0; p < styles.paintStyles.length; p++) {
        if (styles.paintStyles[p].name === styleName) {
          docState.snapshot.paint.set(styleName, JSON.stringify(styles.paintStyles[p]));
          break;
        }
      }
    } else if (textNames.has(styleName)) {
      docState.rows.set("text:" + styleName, child.id);
      for (var t = 0; t < styles.textStyles.length; t++) {
        if (styles.textStyles[t].name === styleName) {
          docState.snapshot.text.set(styleName, JSON.stringify(styles.textStyles[t]));
          break;
        }
      }
    } else if (effectNames.has(styleName)) {
      docState.rows.set("effect:" + styleName, child.id);
      for (var e = 0; e < styles.effectStyles.length; e++) {
        if (styles.effectStyles[e].name === styleName) {
          docState.snapshot.effect.set(styleName, JSON.stringify(styles.effectStyles[e]));
          break;
        }
      }
    }
  }
}

// Initial send
sendAllData();

// Detect existing docs on startup
detectExistingDocs();

// ═══════════════════════════════════════════════════════════════════════
// ─── Audit: broken links & deprecated components (v1.4.1) ───
// Read-only. Scans INSTANCE nodes, groups them by main component and flags:
//   deprecated → the main component lives in a /graveyard/i page, or its
//                (set) description carries "@deprecated", or its name says
//                DEPRECATED
//   missing    → the main component is gone (removed / not found)
//   library    → remote components are verified by the UI through the Figma
//                REST API only (published list + Graveyard page per library).
//                Nothing is imported into the file: importing a component by
//                key would register "imported components" in the document,
//                which shows up as a change in version history — so it is NOT done.
// Performance (v1.4.1): one SYNC mainComponent access per instance, no parent
// walks during the scan — layer paths and nesting are resolved afterwards and
// only for instances of flagged components (typically ~1% of the file). OK
// components travel to the UI as a count, not as thousands of instance rows.
// Nothing here edits the document — only selection/viewport on "Locate".
// ═══════════════════════════════════════════════════════════════════════
var _auditComps = null; // last scan's component records (with light refs) for auditExpand
var AUDIT_GRAVEYARD_RE = /graveyard/i;
var AUDIT_TAG_RE = /@deprecated/i;
var AUDIT_NAME_RE = /\bdeprecated\b/i;
var _auditRunning = false;
var _auditCancel = false;

// "@deprecated → use Button/Primary" | "@deprecated -> Button/Primary" |
// "@deprecated: use X instead" → "Button/Primary"
function auditParseReplacement(desc) {
  if (!desc) return "";
  var lines = String(desc).split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var m = /@deprecated\s*(?:→|->|:|—|–|-)?\s*(?:use\s+)?(.+?)\s*(?:\binstead\b)?\s*\.?\s*$/i.exec(lines[i]);
    if (m && m[1]) return m[1].trim();
  }
  return "";
}

function auditPageOf(node) {
  var p = node;
  try { while (p && p.type !== "PAGE") p = p.parent; } catch (e) { return null; }
  return p && p.type === "PAGE" ? p : null;
}

// Ancestor names (page excluded), outermost INSTANCE ancestor and outermost
// COMPONENT/COMPONENT_SET ancestor, in ONE parent walk.
// The component ancestor is what separates "someone used a retired component in a
// design" from "a live, published component is BUILT ON a retired one" — the second
// propagates to everybody who places that component, so it ranks far higher.
function auditAncestry(node) {
  var parts = [], outer = null, comp = null;
  try {
    var p = node.parent;
    while (p && p.type !== "PAGE" && p.type !== "DOCUMENT") {
      parts.unshift(p.name);
      if (p.type === "INSTANCE") outer = p;
      if (p.type === "COMPONENT" || p.type === "COMPONENT_SET") comp = p; // keep the outermost
      p = p.parent;
    }
  } catch (e) {}
  return { path: parts, outer: outer, comp: comp };
}

// Sync first (legacy access mode — cheap), async only if the sync getter is unavailable
async function auditMainOf(inst) {
  try { return inst.mainComponent; } catch (e) {}
  try { if (typeof inst.getMainComponentAsync === "function") return await inst.getMainComponentAsync(); } catch (e) {}
  return null;
}

// Component record shared by all instances of the same main component
function auditDescribe(mc, inst) {
  var rec = {
    id: "", key: "", name: "", setName: "", variantName: "", remote: false, description: "",
    localPage: "", libraryName: "", status: "ok", detectedBy: [], replacement: "", published: null,
    count: 0, refs: [], instances: undefined,
  };
  if (!mc) { rec.name = inst.name; rec.status = "missing"; rec.detectedBy.push("main component not found"); return rec; }
  var removed = false; try { removed = !!mc.removed; } catch (e) { removed = true; }
  if (removed) { rec.name = inst.name; rec.status = "missing"; rec.detectedBy.push("main component deleted"); return rec; }
  rec.id = mc.id;
  try { rec.key = mc.key || ""; } catch (e) {}
  try { rec.remote = !!mc.remote; } catch (e) {}
  var set = null;
  try { set = mc.parent && mc.parent.type === "COMPONENT_SET" ? mc.parent : null; } catch (e) {}
  rec.setName = set ? set.name : "";
  rec.variantName = set ? mc.name : "";
  rec.name = set ? set.name + " / " + mc.name : mc.name;
  var setDesc = ""; try { setDesc = (set && set.description) || ""; } catch (e) {}
  var mcDesc = ""; try { mcDesc = mc.description || ""; } catch (e) {}
  rec.description = (setDesc + "\n" + mcDesc).trim();

  if (!rec.remote) {
    var pg = auditPageOf(set || mc);
    rec.localPage = pg ? pg.name : "";
    if (pg && AUDIT_GRAVEYARD_RE.test(pg.name)) { rec.status = "deprecated"; rec.detectedBy.push("Graveyard page"); }
  }
  if (AUDIT_TAG_RE.test(rec.description)) {
    rec.status = "deprecated"; rec.detectedBy.push("@deprecated tag");
    rec.replacement = auditParseReplacement(rec.description);
  }
  if (AUDIT_NAME_RE.test(rec.name)) { rec.status = "deprecated"; rec.detectedBy.push("name says DEPRECATED"); }
  return rec;
}

// Progress with live counters so the UI can show what is happening while it waits
function auditProgress(stage, current, total, label, comps, scanned) {
  var keys = Object.keys(comps || {});
  var issueInstances = 0;
  for (var i = 0; i < keys.length; i++) {
    var c = comps[keys[i]];
    if (c.status === "deprecated" || c.status === "missing") issueInstances += c.count;
  }
  figma.ui.postMessage({
    type: "audit-progress", stage: stage, current: current, total: total, label: label || "",
    scanned: scanned || 0, components: keys.length, issues: issueInstances,
  });
}

function auditYield() { return new Promise(function (r) { setTimeout(r, 0); }); }

// Turn the light refs [id, layerName, pageName] of a flagged component into full rows
function auditExpandRefs(rec) {
  var out = [];
  for (var i = 0; i < rec.refs.length; i++) {
    var ref = rec.refs[i];
    var node = null; try { node = figma.getNodeById(ref[0]); } catch (e) {}
    var anc = node ? auditAncestry(node) : { path: [], outer: null };
    out.push({ nodeId: ref[0], layerName: ref[1], pageName: ref[2], path: anc.path,
      nested: !!anc.outer, nestedIn: anc.outer ? anc.outer.name : "",
      inComponent: !!anc.comp, componentName: anc.comp ? anc.comp.name : "" });
  }
  return out;
}

async function runAudit(scope) {
  if (_auditRunning) return;
  _auditRunning = true; _auditCancel = false;
  var t0 = Date.now();
  try {
    var pages = scope === "page" ? [figma.currentPage] : figma.root.children.slice();
    var comps = {};
    var scanned = 0;

    for (var pi = 0; pi < pages.length; pi++) {
      if (_auditCancel) { figma.ui.postMessage({ type: "audit-cancelled" }); return; }
      var page = pages[pi];
      var pageName = page.name;
      auditProgress("scan", pi, pages.length, pageName, comps, scanned);
      var instances = typeof page.findAllWithCriteria === "function"
        ? page.findAllWithCriteria({ types: ["INSTANCE"] })
        : page.findAll(function (n) { return n.type === "INSTANCE"; });

      for (var ii = 0; ii < instances.length; ii++) {
        var inst = instances[ii];
        scanned++;
        var mc = await auditMainOf(inst);
        var gk = mc ? mc.id : "missing:" + inst.name;
        var rec = comps[gk];
        if (!rec) { rec = auditDescribe(mc, inst); comps[gk] = rec; }
        rec.count++;
        rec.refs.push([inst.id, inst.name, pageName]);
        if (scanned % 2000 === 0) {
          auditProgress("scan", pi, pages.length, pageName, comps, scanned);
          await auditYield();
          if (_auditCancel) { figma.ui.postMessage({ type: "audit-cancelled" }); return; }
        }
      }
      await auditYield();
    }
    auditProgress("scan", pages.length, pages.length, "", comps, scanned);

    var list = Object.keys(comps).map(function (k) { return comps[k]; });
    _auditComps = {};
    for (var li = 0; li < list.length; li++) _auditComps[list[li].id || ("missing:" + list[li].name)] = list[li];

    // Full rows (layer path, nesting) now for locally flagged components; remote ones are
    // verified by the UI via REST and expanded on request (audit-expand)
    var payloadList = [];
    for (var pi2 = 0; pi2 < list.length; pi2++) {
      var c2 = list[pi2];
      var copy = {};
      for (var k2 in c2) if (k2 !== "refs" && k2 !== "instances") copy[k2] = c2[k2];
      copy.uid = c2.id || ("missing:" + c2.name);
      if (c2.status !== "ok") copy.instances = auditExpandRefs(c2);
      payloadList.push(copy);
    }

    var fileKey = ""; try { fileKey = figma.fileKey || ""; } catch (e) {}
    figma.ui.postMessage({
      type: "audit-result",
      payload: {
        fileName: figma.root.name, fileKey: fileKey, scope: scope, pages: pages.length,
        instancesScanned: scanned, components: payloadList, durationMs: Date.now() - t0, scannedAt: new Date().toISOString(),
      },
    });
  } catch (err) {
    console.error("[Audit] Error:", err);
    figma.ui.postMessage({ type: "audit-error", message: String((err && err.message) || err) });
  } finally {
    _auditRunning = false;
  }
}

function auditLocate(nodeId) {
  var n = null;
  try { n = figma.getNodeById(nodeId); } catch (e) {}
  if (!n) { figma.notify("That layer no longer exists"); return; }
  var pg = auditPageOf(n);
  if (pg && figma.currentPage !== pg) figma.currentPage = pg;
  figma.currentPage.selection = [n];
  figma.viewport.scrollAndZoomIntoView([n]);
}

// UI asks for full rows of components it flagged after the REST check
function auditExpand(uids) {
  var rows = {};
  if (_auditComps) {
    for (var i = 0; i < uids.length; i++) {
      var rec = _auditComps[uids[i]];
      if (rec && rec.refs) rows[uids[i]] = auditExpandRefs(rec);
    }
  }
  figma.ui.postMessage({ type: "audit-expanded", rows: rows });
}

// Cheap facts about the file so the UI can warn before a long scan
function auditFileInfo() {
  var fileKey = ""; try { fileKey = figma.fileKey || ""; } catch (e) {}
  var topLevel = 0;
  try { for (var i = 0; i < figma.root.children.length; i++) topLevel += figma.root.children[i].children.length; } catch (e) {}
  figma.ui.postMessage({ type: "audit-file-info", payload: { fileName: figma.root.name, fileKey: fileKey, pages: figma.root.children.length, topLevelNodes: topLevel } });
}

// Last scan per file, kept in clientStorage (per user); newest 3 files only
function auditLastKey() { var k = ""; try { k = figma.fileKey || ""; } catch (e) {} return k || figma.root.name; }
async function auditGetLast() {
  var all = null; try { all = await figma.clientStorage.getAsync("audit-last"); } catch (e) {}
  var mine = all && all[auditLastKey()] ? all[auditLastKey()] : null;
  figma.ui.postMessage({ type: "audit-last-data", payload: mine });
}
async function auditSaveLast(data) {
  try {
    var all = (await figma.clientStorage.getAsync("audit-last")) || {};
    all[auditLastKey()] = data;
    var keys = Object.keys(all).sort(function (a, b) { return (all[b].savedAt || 0) - (all[a].savedAt || 0); });
    for (var i = 3; i < keys.length; i++) delete all[keys[i]];
    await figma.clientStorage.setAsync("audit-last", all);
    figma.ui.postMessage({ type: "audit-last-saved", ok: true });
  } catch (e) {
    console.warn("[Audit] could not save last scan:", e);
    figma.ui.postMessage({ type: "audit-last-saved", ok: false });
  }
}

// ─── Audit report on canvas ───────────────────────────────────────────
// Builds a documentation frame from the report the panel is showing, so the
// canvas always matches the screen. Unlike the scan itself this DOES create
// content: it is an explicit action, never automatic.
// Status colours verified on white at 16px: amber #E0A800 with #1A1A1A text
// 8.1:1 (AAA); red #C4341B, green #157A44 and purple #6B4FD6 with white text
// 5.0:1, 5.4:1 and 5.7:1 (all AA).

async function dAuditStat(parent, label, value, pctText, colour) {
  var card = df("Stat — " + label, "VERTICAL", 8);
  card.paddingLeft = 20; card.paddingRight = 20; card.paddingTop = 18; card.paddingBottom = 18;
  card.cornerRadius = 10;
  card.fills = [{ type: "SOLID", color: DC.white }];
  card.strokes = [{ type: "SOLID", color: DC.divider }]; card.strokeWeight = 1;

  var top = df("Label", "HORIZONTAL", 8);
  top.counterAxisAlignItems = "CENTER";
  var dot = figma.createEllipse();
  dot.name = "Dot"; dot.resize(12, 12);
  dot.fills = [{ type: "SOLID", color: colour }];
  top.appendChild(dot);
  ac(top, await dt(label, DC.MIN, "Medium", DC.textSec));
  ac(card, top, "FILL", "HUG");

  ac(card, await dt(String(value), 34, "Bold", DC.text), "FILL", "HUG");
  ac(card, await dt(pctText, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
  ac(parent, card, "FILL", "HUG");
  return card;
}

async function dAuditHealthBar(parent, r) {
  var wrap = df("Health", "VERTICAL", 10);
  wrap.paddingLeft = DC.P; wrap.paddingRight = DC.P; wrap.paddingTop = 4; wrap.paddingBottom = 20;

  var inner = DC.W - DC.P * 2;
  var total = Math.max(1, r.instancesScanned);
  var segs = [
    ["Linked", r.counts.linked, DC.auOk],
    ["Broken", r.counts.broken, DC.auMiss],
    ["In components", r.counts.inComp, DC.auComp],
    ["In designs", r.counts.inDesign, DC.auDep],
    ["Info only", r.counts.info, DC.auInfo],
  ];
  var bar = df("Bar", "HORIZONTAL", 0);
  bar.cornerRadius = 8; bar.clipsContent = true;
  bar.fills = [{ type: "SOLID", color: DC.divider }];
  var used = 0;
  for (var i = 0; i < segs.length; i++) {
    if (!segs[i][1]) continue;
    var w = i === segs.length - 1 ? Math.max(3, inner - used) : Math.max(3, Math.round(inner * segs[i][1] / total));
    used += w;
    var seg = figma.createRectangle();
    seg.name = segs[i][0]; seg.resize(w, 16);
    seg.fills = [{ type: "SOLID", color: segs[i][2] }];
    bar.appendChild(seg);
  }
  ac(wrap, bar, "HUG", "HUG");
  ac(parent, wrap, "FILL", "HUG");
}

async function dAuditSummary(parent, r) {
  var row = df("Summary", "HORIZONTAL", 16);
  row.paddingLeft = DC.P; row.paddingRight = DC.P; row.paddingTop = 28; row.paddingBottom = 16;
  var total = Math.max(1, r.instancesScanned);
  var pc = function (n) { return (100 * n / total).toFixed(1) + "% of all instances"; };
  await dAuditStat(row, "Linked", r.counts.linked.toLocaleString(), pc(r.counts.linked), DC.auOk);
  await dAuditStat(row, "Broken", r.counts.broken.toLocaleString(), pc(r.counts.broken), DC.auMiss);
  await dAuditStat(row, "In components", r.counts.inComp.toLocaleString(), pc(r.counts.inComp), DC.auComp);
  await dAuditStat(row, "In designs", r.counts.inDesign.toLocaleString(), pc(r.counts.inDesign), DC.auDep);
  await dAuditStat(row, "Info only", r.counts.info.toLocaleString(), pc(r.counts.info), DC.auInfo);
  ac(parent, row, "FILL", "HUG");
}

// What each status means — the report has to stand on its own away from the panel
async function dAuditLegend(parent) {
  var items = [
    ["Broken", DC.auMiss, "The component no longer exists, or is no longer published in any reachable library. Figma still draws it from its local cache, so it looks fine, but the link is dead and it will never update again. This is the only category that is genuinely broken."],
    ["In component", DC.auComp, "A live, published component is built on a retired one. Nothing looks wrong, but every design that places this component inherits the retired dependency, so fixing it once here clears it everywhere. Chase these first after broken links."],
    ["In design", DC.auDep, "A retired component used directly in a design or template. Nothing is broken; swap it for the current component next time that screen is touched."],
    ["Info only", DC.auInfo, "Cannot be fixed where it sits: nested inside another instance, so it is fixed in the component that contains it, or sitting on a Graveyard page, which is retired material and expected."],
    ["Linked", DC.auOk, "Points to a component that exists, is published and has not been retired. Nothing to do."],
  ];
  await dSectionTitle(parent, "What the statuses mean");
  var box = df("Legend", "VERTICAL", 14);
  box.paddingLeft = DC.P; box.paddingRight = DC.P; box.paddingBottom = 8;
  for (var i = 0; i < items.length; i++) {
    var line = df("Item", "HORIZONTAL", 14);
    line.counterAxisAlignItems = "MIN";
    var b = await dBadge(items[i][0], items[i][1], items[i][0] === "In design" ? DC.auDepTxt : DC.white);
    b.cornerRadius = 6;
    ac(line, b);
    var txt = df("Text", "VERTICAL", 0);
    ac(txt, await dt(items[i][2], DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
    ac(line, txt, "FILL", "HUG");
    ac(box, line, "FILL", "HUG");
  }
  ac(parent, box, "FILL", "HUG");
}

async function dAuditByPage(parent, r) {
  if (!r.byPage || !r.byPage.length) return;
  await dSectionTitle(parent, "Issues by live page");
  var box = df("By page", "VERTICAL", 10);
  box.paddingLeft = DC.P; box.paddingRight = DC.P; box.paddingBottom = 12;
  ac(box, await dt("Pages the team works on today. Graveyard and other retired areas are excluded \u2014 they appear under Info only.", DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
  var max = r.byPage[0][1] || 1;
  for (var i = 0; i < r.byPage.length; i++) {
    var line = df("Page", "HORIZONTAL", 16);
    line.counterAxisAlignItems = "CENTER";
    var nameCol = df("Name", "VERTICAL", 0);
    ac(nameCol, await dt(r.byPage[i][0], DC.MIN, "Medium", DC.text), "FILL", "HUG");
    ac(line, nameCol, "FIXED", "HUG", 340);

    var track = df("Track", "HORIZONTAL", 0);
    track.cornerRadius = 5; track.clipsContent = true;
    track.fills = [{ type: "SOLID", color: DC.divider }];
    var fill = figma.createRectangle();
    fill.name = "Fill";
    fill.resize(Math.max(4, Math.round(620 * r.byPage[i][1] / max)), 10);
    fill.fills = [{ type: "SOLID", color: DC.auMiss }];
    track.appendChild(fill);
    ac(line, track, "FILL", "HUG");

    var cnt = df("Count", "VERTICAL", 0);
    cnt.counterAxisAlignItems = "MAX";
    ac(cnt, await dt(String(r.byPage[i][1]), DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");
    ac(line, cnt, "FIXED", "HUG", 60);
    ac(box, line, "FILL", "HUG");
  }
  ac(parent, box, "FILL", "HUG");
}

async function dAuditTableCell(row, text, weight, colour, mode, width) {
  var cell = df("Cell", "VERTICAL", 2);
  cell.paddingLeft = 14; cell.paddingRight = 14; cell.paddingTop = 12; cell.paddingBottom = 12;
  ac(cell, await dt(text || "—", DC.MIN, weight, colour), "FILL", "HUG");
  ac(row, cell, mode, "FILL", width);
  return cell;
}

async function dAuditIssues(parent, r) {
  await dSectionTitle(parent, "What to fix — " + r.issueCount.toLocaleString() + " instance" + (r.issueCount === 1 ? "" : "s") + " in " + r.rows.length + " component" + (r.rows.length === 1 ? "" : "s"));
  var wrap = df("Issues", "VERTICAL", 0);
  wrap.paddingLeft = DC.P; wrap.paddingRight = DC.P; wrap.paddingBottom = 16;

  var tbl = df("Table", "VERTICAL", 0);
  tbl.cornerRadius = 8; tbl.clipsContent = true;
  tbl.strokes = [{ type: "SOLID", color: DC.divider }]; tbl.strokeWeight = 1;

  var hdr = df("Header", "HORIZONTAL", 0);
  hdr.fills = [{ type: "SOLID", color: DC.badgeBg }];
  await dAuditTableCell(hdr, "Status", "Semi Bold", DC.text, "FIXED", 130);
  await dAuditTableCell(hdr, "Component", "Semi Bold", DC.text, "FILL");
  await dAuditTableCell(hdr, "Page", "Semi Bold", DC.text, "FIXED", 230);
  await dAuditTableCell(hdr, "Location", "Semi Bold", DC.text, "FILL");
  ac(tbl, hdr, "FILL", "HUG");

  var shown = r.rows.slice(0, 300);
  for (var i = 0; i < shown.length; i++) {
    var it = shown[i];
    var line = figma.createRectangle();
    line.name = "Divider"; line.resize(100, 1);
    line.fills = [{ type: "SOLID", color: DC.divider }];
    ac(tbl, line, "FILL", "FIXED");

    var tr = df("Row — " + it.name, "HORIZONTAL", 0);
    tr.fills = [{ type: "SOLID", color: DC.white }];

    var sevMap = { broken: ["Broken", DC.auMiss, DC.white], component: ["In component", DC.auComp, DC.white], design: ["In design", DC.auDep, DC.auDepTxt] };
    var sev = sevMap[it.sev] || sevMap.design;
    var stCell = df("Cell", "VERTICAL", 0);
    stCell.paddingLeft = 14; stCell.paddingRight = 14; stCell.paddingTop = 12; stCell.paddingBottom = 12;
    var bdg = await dBadge(sev[0], sev[1], sev[2]);
    bdg.cornerRadius = 5;
    ac(stCell, bdg);
    ac(tr, stCell, "FIXED", "FILL", 130);

    var nameCell = df("Cell", "VERTICAL", 3);
    nameCell.paddingLeft = 14; nameCell.paddingRight = 14; nameCell.paddingTop = 12; nameCell.paddingBottom = 12;
    ac(nameCell, await dt(it.name + (it.count > 1 ? "  ×" + it.count : ""), DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");
    var sub = (it.sev === "component" && it.builtInto ? "built into " + it.builtInto + "  ·  " : "") + (it.detectedBy || "");
    if (it.replacement) sub += "  ·  replace with: " + it.replacement;
    if (it.library) sub += "  ·  " + it.library;
    ac(nameCell, await dt(sub, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
    ac(tr, nameCell, "FILL", "FILL");

    await dAuditTableCell(tr, it.page, "Regular", DC.text, "FIXED", 230);

    var locCell = df("Cell", "VERTICAL", 3);
    locCell.paddingLeft = 14; locCell.paddingRight = 14; locCell.paddingTop = 12; locCell.paddingBottom = 12;
    ac(locCell, await dt(it.frame || "—", DC.MIN, "Medium", DC.text), "FILL", "HUG");
    if (it.path) ac(locCell, await dt(it.path, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
    ac(tr, locCell, "FILL", "FILL");

    ac(tbl, tr, "FILL", "HUG");
  }
  ac(wrap, tbl, "FILL", "HUG");

  if (r.rows.length > 300) {
    var more = df("More", "VERTICAL", 0);
    more.paddingTop = 12;
    ac(more, await dt("Showing the first 300 components. The full list is in the CSV export.", DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
    ac(wrap, more, "FILL", "HUG");
  }
  ac(parent, wrap, "FILL", "HUG");
}

async function dAuditInfoNote(parent, r) {
  if (!r.counts.info) return;
  await dSectionTitle(parent, "Info only — " + r.counts.info.toLocaleString() + " instances");
  var box = df("Info note", "VERTICAL", 0);
  box.paddingLeft = DC.P; box.paddingRight = DC.P; box.paddingBottom = 16;
  var parts = [];
  if (r.info.nested) parts.push(r.info.nested.toLocaleString() + " nested inside another component");
  if (r.info.graveyard) parts.push(r.info.graveyard.toLocaleString() + " sitting on a Graveyard page");
  var txt = parts.join(" and ") + ". These are not counted as issues: a nested instance is fixed once in the "
    + "library and every copy is fixed with it, and retired material using other retired material is expected.";
  ac(box, await dDescCallout(txt), "FILL", "HUG");
  ac(parent, box, "FILL", "HUG");
}

// Private components: skipped on purpose, and explicitly harmless
async function dAuditPrivateNote(parent, r) {
  if (!r.privateHelpers) return;
  await dSectionTitle(parent, "Private components — " + r.privateHelpers.toLocaleString() + " skipped");
  var box = df("Private note", "VERTICAL", 0);
  box.paddingLeft = DC.P; box.paddingRight = DC.P; box.paddingBottom = 16;
  var txt = "A component whose name starts with a dot or an underscore is private in Figma: it is never "
    + "published to a library, because it only exists as an internal part of a published component. The "
    + "library API has nothing to say about them, so they are skipped rather than checked. They count as "
    + "correctly linked and do not affect the health score above.";
  ac(box, await dDescCallout(txt), "FILL", "HUG");
  ac(parent, box, "FILL", "HUG");
}

async function dAuditHeader(parent, r) {
  var h = df("Header", "VERTICAL", 0);
  h.fills = [{ type: "SOLID", color: DC.headerBg }];
  h.paddingLeft = DC.P; h.paddingRight = DC.P; h.paddingTop = DC.P; h.paddingBottom = DC.P;

  // Title on the left, health score on the right. Laid out with auto-layout rather
  // than absolute positioning, so the score can never be clipped by a fixed width.
  var top = df("Top", "HORIZONTAL", 40);
  top.counterAxisAlignItems = "CENTER";

  var left = df("Title", "VERTICAL", 12);
  ac(left, await dt("Component Audit", 42, "Bold", DC.white), "FILL", "HUG");

  var bits = [];
  if (r.counts.broken) bits.push(r.counts.broken.toLocaleString() + " broken");
  if (r.counts.inComp) bits.push(r.counts.inComp.toLocaleString() + " in components");
  if (r.counts.inDesign) bits.push(r.counts.inDesign.toLocaleString() + " in designs");
  if (!bits.length) bits.push("nothing to fix");
  var counts = bits.join(" \u00b7 ") + "  \u00b7  " + r.instancesScanned.toLocaleString()
    + " instances scanned \u00b7 " + (r.scope === "page" ? "current page" : r.pages + " pages");
  var countLine = await dt(counts.toUpperCase(), DC.MIN, "Medium", DC.white);
  countLine.opacity = 0.5;
  countLine.letterSpacing = { value: 2, unit: "PIXELS" };
  ac(left, countLine, "FILL", "HUG");

  var sub = await dt(r.fileName + (r.libraries && r.libraries.length ? "  \u00b7  libraries verified: " + r.libraries.join(", ") : ""), DC.MIN, "Regular", DC.white);
  sub.opacity = 0.4;
  ac(left, sub, "FILL", "HUG");
  ac(top, left, "FILL", "HUG");

  // Health score
  var pctNum = r.health != null ? r.health : 100;
  var col = pctNum >= 99.5 ? DC.hpGood : (pctNum >= 98 ? DC.hpWarn : DC.hpBad);
  var score = df("Health score", "VERTICAL", 4);
  score.counterAxisAlignItems = "MAX";

  ac(score, await dt(pctNum.toFixed(2).replace(/\.00$/, "") + "%", 64, "Bold", col));
  var lbl = await dt("DESIGN SYSTEM HEALTH", DC.MIN, "Medium", DC.white);
  lbl.opacity = 0.55; lbl.letterSpacing = { value: 2, unit: "PIXELS" };
  ac(score, lbl);
  if (r.componentsTotal) {
    var camo = await dt(r.componentsAffected.toLocaleString() + " of " + r.componentsTotal.toLocaleString() + " components affected", DC.MIN, "Regular", DC.white);
    camo.opacity = 0.4;
    ac(score, camo);
  }
  // Force HUG on both axes so the parent can never squeeze it into a fixed width
  score.primaryAxisSizingMode = "AUTO";
  score.counterAxisSizingMode = "AUTO";
  score.clipsContent = false;
  ac(top, score, "HUG", "HUG");

  ac(h, top, "FILL", "HUG");
  ac(parent, h, "FILL", "HUG");
}

async function generateAuditDocumentation(r) {
  try {
    figma.ui.postMessage({ type: "audit-doc-progress", stage: "start" });
    _lf.clear();
    await lf("Inter", "Regular"); await lf("Inter", "Medium");
    await lf("Inter", "Semi Bold"); await lf("Inter", "Bold");

    var page = figma.currentPage;
    var startX = 0;
    for (var n = 0; n < page.children.length; n++) {
      var right = page.children[n].x + page.children[n].width;
      if (right > startX) startX = right;
    }
    startX += 200;

    var main = figma.createFrame();
    main.name = "Component Audit — " + r.fileName;
    main.layoutMode = "VERTICAL";
    main.counterAxisSizingMode = "FIXED";
    main.primaryAxisSizingMode = "AUTO";
    main.resize(DC.W, 100);
    main.fills = [{ type: "SOLID", color: DC.white }];
    main.itemSpacing = 0;
    page.appendChild(main);
    main.x = startX; main.y = 0;

    await dAuditHeader(main, r);
    await dAuditSummary(main, r);
    await dAuditHealthBar(main, r);
    dDiv(main);
    await dAuditByPage(main, r);
    dDiv(main);
    if (r.rows.length) { await dAuditIssues(main, r); dDiv(main); }
    await dAuditInfoNote(main, r);
    await dAuditPrivateNote(main, r);
    dDiv(main);
    await dAuditLegend(main);
    dDiv(main);
    await dFooter(main);

    figma.viewport.scrollAndZoomIntoView([main]);
    figma.currentPage.selection = [main];
    figma.ui.postMessage({ type: "audit-doc-progress", stage: "done" });
    figma.notify("Audit report created on canvas");
  } catch (err) {
    console.error("[Audit doc] Error:", err);
    figma.ui.postMessage({ type: "audit-doc-progress", stage: "error", message: String((err && err.message) || err) });
    figma.notify("Could not create the report: " + String((err && err.message) || err));
  }
}

// ─── Listen for UI messages ───

figma.ui.onmessage = async (msg) => {
  if (msg.type === "sync") await sendAllData();
  if (msg.type === "close") figma.closePlugin();
  if (msg.type === "generate-docs") await generateDocumentation(msg.payload);
  if (msg.type === "generate-var-docs") await generateVariableDocumentation(msg.payload);
  if (msg.type === "resize-ui") figma.ui.resize(msg.width, msg.height);

  // ─── Audit (read-only) ───
  if (msg.type === "audit-scan") await runAudit(msg.scope || "all");
  if (msg.type === "audit-locate") auditLocate(msg.nodeId);
  if (msg.type === "audit-cancel") _auditCancel = true;
  if (msg.type === "audit-expand") auditExpand(msg.uids || []);

  // ─── Export naming preferences ───
  if (msg.type === "naming-get") {
    var nOpts = await figma.clientStorage.getAsync("naming-opts");
    figma.ui.postMessage({ type: "naming-data", payload: nOpts || null });
  }
  if (msg.type === "naming-save") await figma.clientStorage.setAsync("naming-opts", msg.payload);
  if (msg.type === "audit-generate-doc") await generateAuditDocumentation(msg.payload);
  if (msg.type === "audit-file-info") auditFileInfo();
  if (msg.type === "audit-get-last") await auditGetLast();
  if (msg.type === "audit-save-last") await auditSaveLast(msg.payload);
  if (msg.type === "audit-get-token") {
    var auTok = await figma.clientStorage.getAsync("audit-token");
    figma.ui.postMessage({ type: "audit-token-data", payload: auTok || null });
  }
  if (msg.type === "audit-save-token") {
    await figma.clientStorage.setAsync("audit-token", msg.payload);
    figma.ui.postMessage({ type: "audit-token-data", payload: msg.payload });
  }
  if (msg.type === "audit-clear-token") {
    await figma.clientStorage.deleteAsync("audit-token");
    figma.ui.postMessage({ type: "audit-token-data", payload: null });
  }

  // ─── Bitbucket credential storage ───
  if (msg.type === "bb-get-creds") {
    var creds = await figma.clientStorage.getAsync("bb-creds");
    figma.ui.postMessage({ type: "bb-creds-data", payload: creds || null });
  }
  if (msg.type === "bb-save-creds") {
    await figma.clientStorage.setAsync("bb-creds", msg.payload);
    figma.ui.postMessage({ type: "bb-creds-saved" });
  }
  if (msg.type === "bb-delete-creds") {
    await figma.clientStorage.deleteAsync("bb-creds");
    figma.ui.postMessage({ type: "bb-creds-deleted" });
  }

  // ─── Saved reviewers storage ───
  if (msg.type === "bb-get-reviewers") {
    var reviewers = await figma.clientStorage.getAsync("bb-saved-reviewers");
    figma.ui.postMessage({ type: "bb-reviewers-data", payload: reviewers || [] });
  }
  if (msg.type === "bb-save-reviewers") {
    await figma.clientStorage.setAsync("bb-saved-reviewers", msg.payload);
    figma.ui.postMessage({ type: "bb-reviewers-saved" });
  }

  // ─── Last PR config storage ───
  if (msg.type === "bb-get-config") {
    var config = await figma.clientStorage.getAsync("bb-last-config");
    figma.ui.postMessage({ type: "bb-config-data", payload: config || null });
  }
  if (msg.type === "bb-save-config") {
    await figma.clientStorage.setAsync("bb-last-config", msg.payload);
  }
};
