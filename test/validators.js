// Structural validators: does each export actually hold together as a file?
"use strict";
const assert = require("node:assert/strict");

// ── CSS: every declaration well-formed, no duplicate names per block, every var() resolvable ──
function validateCss(css) {
  const noComments = css.replace(/\/\*[\s\S]*?\*\//g, "");
  const blocks = [...noComments.matchAll(/([^{}]+)\{([^{}]*)\}/g)].map((m) => ({ selector: m[1].trim(), body: m[2] }));
  assert.ok(blocks.length >= 1, "CSS has at least one rule block");
  const problems = [];
  for (const b of blocks) {
    const decls = b.body.split("\n").map((l) => l.trim()).filter((l) => l.startsWith("--"));
    const seen = new Map();
    const defined = new Set();
    for (const d of decls) {
      const m = /^(--[a-z0-9-]+)\s*:\s*(.+);\s*(\/\*.*\*\/)?$/i.exec(d);
      if (!m) { problems.push(`malformed: ${d}`); continue; }
      if (!/^--[a-z0-9]([a-z0-9-]*[a-z0-9])?$/i.test(m[1])) problems.push(`bad property name: ${m[1]}`);
      if (seen.has(m[1])) problems.push(`duplicate in "${b.selector}": ${m[1]} (${seen.get(m[1])} vs ${m[2]})`);
      seen.set(m[1], m[2]); defined.add(m[1]);
    }
    for (const d of decls) {
      for (const ref of d.matchAll(/var\((--[a-z0-9-]+)\)/gi)) {
        if (!defined.has(ref[1])) problems.push(`unresolved reference in "${b.selector}": ${ref[1]}`);
      }
    }
  }
  assert.deepEqual(problems, [], "CSS structural problems");
  return blocks;
}

// ── DTCG: valid JSON, every leaf has $type+$value, no "alias" type, every {ref} resolves ──
function walkTokens(node, path, out) {
  if (node && typeof node === "object" && "$value" in node) { out.push({ path, token: node }); return; }
  if (node && typeof node === "object") for (const k of Object.keys(node)) if (!k.startsWith("$")) walkTokens(node[k], [...path, k], out);
}
function resolvePath(root, dotted) {
  return dotted.split(".").reduce((n, k) => (n && typeof n === "object" ? n[k] : undefined), root);
}
function validateDtcg(json, { modeRoots = false } = {}) {
  const t = JSON.parse(json);
  const roots = modeRoots ? Object.values(t) : [t];
  const problems = [];
  let count = 0;
  for (const root of roots) {
    const leaves = []; walkTokens(root, [], leaves);
    for (const { path, token } of leaves) {
      count++;
      if (!("$type" in token)) problems.push(`missing $type at ${path.join(".")}`);
      if (token.$type === "alias") problems.push(`$type "alias" is not a DTCG type at ${path.join(".")}`);
      if (typeof token.$value === "string") {
        for (const ref of token.$value.matchAll(/\{([^}]+)\}/g)) {
          const target = resolvePath(root, ref[1]);
          if (!target || !("$value" in target)) problems.push(`unresolved reference ${ref[0]} at ${path.join(".")}`);
          if (/\//.test(ref[1])) problems.push(`slash in reference ${ref[0]} at ${path.join(".")} (must be dot-separated)`);
        }
      }
    }
  }
  assert.ok(count > 0, "DTCG has at least one token");
  assert.deepEqual(problems, [], "DTCG structural problems");
  return t;
}

// ── Dart: unique class names, valid identifiers, private constructor, typed constants ──
function validateDart(dart) {
  const classes = [...dart.matchAll(/abstract class (\w+)\s*\{/g)].map((m) => m[1]);
  assert.ok(classes.length >= 1, "Dart has at least one class");
  const dup = classes.filter((c, i) => classes.indexOf(c) !== i);
  assert.deepEqual(dup, [], "duplicate Dart class names");
  const problems = [];
  for (const c of classes) {
    if (!/^[A-Z][A-Za-z0-9]*$/.test(c)) problems.push(`class name not PascalCase identifier: ${c}`);
    if (!dart.includes(`${c}._();`)) problems.push(`class ${c} lacks private constructor`);
  }
  for (const m of dart.matchAll(/static const (?:(\w+) )?(\w+) = ([^;\n]+)/g)) {
    const [, type, name, value] = m;
    if (type && !["Color", "double", "String", "bool", "int", "TextStyle", "LinearGradient", "RadialGradient", "SweepGradient", "List"].includes(type)) problems.push(`unexpected Dart type ${type} for ${name}`);
    if (!/^[a-z_][A-Za-z0-9_]*$/.test(name)) problems.push(`constant not a valid identifier: ${name}`);
    // a bare number with no declared type would be inferred as int by Dart
    if (!type && /^-?\d+$/.test(value.trim())) problems.push(`untyped integer literal ${name} = ${value} (Dart infers int; declare double)`);
  }
  // constants must be unique within a class
  for (const cls of dart.split(/abstract class /).slice(1)) {
    const names = [...cls.matchAll(/static const (?:\w+ )?(\w+) =/g)].map((m) => m[1]);
    const d = names.filter((n, i) => names.indexOf(n) !== i);
    if (d.length) problems.push(`duplicate constants in ${cls.split(/\s/)[0]}: ${[...new Set(d)].join(", ")}`);
  }
  assert.deepEqual(problems, [], "Dart structural problems");
  return classes;
}

module.exports = { validateCss, validateDtcg, validateDart };
