// DS Styles & Variables Extractor v2.0
// Extracts Styles (Paint, Text, Effect, Grid)
// Extracts Variables (all Collections, all Modes, resolved values)
// Supports live SYNC + multi-format output (JSON, CSS, Flutter, W3C DTCG)

figma.showUI(__html__, { width: 640, height: 760, themeColors: true });

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
                    // Unwrap alias to get the final value
                    if (resolved.type === 'alias' && resolved.resolvedValue) {
                      boundVarModes[prop][mode.name] = resolved.resolvedValue;
                    } else if (resolved.type !== 'alias') {
                      boundVarModes[prop][mode.name] = resolved;
                    }
                  }
                }
                // Capture mode names from the first multi-mode collection
                if (!boundVarModeNames) {
                  boundVarModeNames = col.modes.map(function(m) { return m.name; });
                }
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

// ─── Send data to UI ───

function sendAllData() {
  const styles = extractAllStyles();
  const variables = extractAllVariables();

  figma.ui.postMessage({
    type: "all-data",
    payload: { styles, variables },
  });
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
  var descCol = df("Description", "VERTICAL", 8);
  var nameParts = style.name.split("/");
  var shortName = nameParts.length > 1 ? nameParts.slice(1).join(" / ") : nameParts[0];
  ac(descCol, await dt(shortName, DC.MIN, "Semi Bold", DC.text), "FILL", "HUG");
  if (style.description) {
    ac(descCol, await dt(style.description, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
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
  await addB(bv.fontStyle ? "$" + bv.fontStyle : "$font-weight-" + style.fontStyle.toLowerCase().replace(/\s+/g, "-"));
  if (style.letterSpacing) {
    var ls = style.letterSpacing.value;
    await addB(bv.letterSpacing ? "$" + bv.letterSpacing : "$letter-spacing-" + (ls === 0 ? "00" : ls.toFixed(1)));
  }
  ac(descCol, badges, "FILL", "HUG");
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
  var descCol = df("Description", "VERTICAL", 4);
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
    ac(descCol, await dt(style.description, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
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
    ac(descCol, await dt(style.description, DC.MIN, "Regular", DC.textSec), "FILL", "HUG");
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

// ─── Listen for UI messages ───

figma.ui.onmessage = async (msg) => {
  if (msg.type === "sync") sendAllData();
  if (msg.type === "close") figma.closePlugin();
  if (msg.type === "generate-docs") await generateDocumentation(msg.payload);
  if (msg.type === "generate-var-docs") await generateVariableDocumentation(msg.payload);
  if (msg.type === "resize-ui") figma.ui.resize(msg.width, msg.height);

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
