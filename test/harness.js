// Loads the REAL plugin UI script (the inline <script> in ui.html) into a stub
// browser, so the tests exercise the same formatters the plugin ships.
// Zero dependencies: node:vm + a permissive DOM proxy.
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const html = fs.readFileSync(path.join(__dirname, "..", "ui.html"), "utf8");
const script = html.slice(html.indexOf("<script>") + 8, html.lastIndexOf("</script>"));

function makeDom() {
  const store = {};
  const mk = (id) => store[id] || (store[id] = new Proxy({
    id, _html: "", style: {}, dataset: {}, textContent: "", value: "", checked: true, className: "",
    classList: { _c: new Set(), add(x) { this._c.add(x); }, remove(x) { this._c.delete(x); },
      toggle(x, f) { if (f === undefined) { this._c.has(x) ? this._c.delete(x) : this._c.add(x); } else { f ? this._c.add(x) : this._c.delete(x); } return this._c.has(x); },
      contains(x) { return this._c.has(x); } },
    setAttribute() {}, getAttribute() { return ""; }, appendChild() {}, removeChild() {}, focus() {}, select() {},
    click() {}, remove() {}, querySelectorAll() { return []; }, querySelector() { return mk("q"); }, closest() { return mk("l"); },
    getBoundingClientRect() { return { left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 }; }, offsetHeight: 40,
  }, {
    get(t, k) { return k === "innerHTML" ? t._html : t[k]; },
    set(t, k, v) { if (k === "innerHTML") t._html = v; else t[k] = v; return true; },
  }));
  return { store, mk };
}

function load() {
  const { store, mk } = makeDom();
  const posted = [];
  const sandbox = {
    console, JSON, Math, Date, Set, Map, Object, Array, String, Number, Boolean, RegExp, Promise, Error, TypeError,
    parseInt, parseFloat, isNaN, encodeURIComponent, decodeURIComponent, btoa: (s) => Buffer.from(s, "binary").toString("base64"),
    document: { getElementById: mk, createElement: () => mk("tmp"), querySelectorAll: () => [], querySelector: (s) => mk("s" + s), body: { appendChild() {}, removeChild() {} }, addEventListener() {}, execCommand() { return true; } },
    window: { innerWidth: 900, innerHeight: 860 },
    parent: { postMessage(m) { posted.push(m && m.pluginMessage); } },
    navigator: {},
    fetch: () => Promise.reject(new Error("no network in tests")),
    URL: { createObjectURL: () => "blob:x", revokeObjectURL() {} },
    Blob: function (parts) { this.parts = parts; },
    setTimeout: (f) => { if (typeof f === "function") f(); return 0; }, clearTimeout() {},
    setInterval: () => 0, clearInterval() {},
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(script, sandbox, { filename: "ui.html<script>" });
  // expose what the tests need
  const expose = (name) => vm.runInContext(`typeof ${name} === 'function' || typeof ${name} !== 'undefined' ? ${name} : undefined`, sandbox);
  const api = {};
  for (const n of ["formatVariablesCSS", "formatVariablesFlutter", "formatVariablesDTCG", "formatVariablesFigma",
                   "formatStylesCSS", "formatStylesFlutter", "formatStylesDTCG",
                   "getFormattedOutput", "getFileInfo", "pubDefaultFilePath", "pubDefaultBranch", "getFlutterClassName",
                   "auFileSlug", "auCurrentFileSlug", "auColNamer", "auTokenSegments", "auCssSlug",
                   "auditModel", "auditRender", "auditCsv", "auditReportPayload", "auSev"]) {
    api[n] = expose(n);
  }
  api.set = (code) => vm.runInContext(code, sandbox);
  api.get = (expr) => vm.runInContext(expr, sandbox);
  api.store = store;
  api.posted = posted;
  // convenience setters
  api.setVariables = (data, modes) => {
    sandbox.__d = data; sandbox.__m = modes;
    api.set(`variablesData = __d; allModes = __m; selectedModes = new Set(__m); selectedCollections = new Set(__d.collections.map(c => c.id)); currentSection = 'variables';`);
  };
  api.setStyles = (data) => { sandbox.__s = data; api.set(`stylesData = __s; currentSection = 'styles';`); };
  api.setFormat = (f) => api.set(`currentFormat = '${f}';`);
  api.setNaming = (o) => { sandbox.__o = o; api.set(`auNameOpts = Object.assign({ collection: true, dedupe: true, prefix: '', aliases: {}, style: 'camel' }, __o);`); };
  api.setFile = (n) => { sandbox.__f = n; api.set(`auFileName = __f;`); };
  api.setAudit = (p) => { sandbox.__p = p; api.set(`auditLast = __p; auditFilter = 'all';`); };
  api.onmessage = (msg) => { sandbox.__msg = msg; api.set(`window.onmessage({ data: { pluginMessage: __msg } });`); };
  return api;
}

module.exports = { load };
