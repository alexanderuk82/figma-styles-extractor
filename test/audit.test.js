// Audit results: severity grading, what counts as an issue, and that the panel
// renders (a blank panel was a real bug once).
"use strict";
const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const { load } = require("./harness");

const inst = (id, page, o = {}) => ({ nodeId: id, layerName: "L", pageName: page, path: ["A"], nested: false, nestedIn: "", inComponent: false, componentName: "", ...o });
function payload() {
  return { fileName: "Acme", fileKey: "k", scope: "all", pages: 3, instancesScanned: 1000, durationMs: 5000, scannedAt: new Date().toISOString(), libraries: ["Lib"], componentsTotal: 4,
    components: [
      { uid: "a", id: "a", key: "k1", name: "Tooltip / Default", setName: "Tooltip", variantName: "Default", remote: true, status: "deprecated", detectedBy: ["Graveyard page"], replacement: "", libraryName: "Lib", count: 4,
        instances: [inst("1", "Screens"), inst("2", "Components", { inComponent: true, componentName: "Card" }), inst("3", "Graveyard"), inst("4", "Screens", { nested: true, nestedIn: "Card" })] },
      { uid: "b", id: "b", key: "k2", name: ".Slot", remote: true, status: "missing", detectedBy: ["gone"], replacement: "", count: 1, instances: [inst("5", "Screens")] },
      { uid: "c", id: "c", key: "k3", name: "Button", remote: true, status: "ok", detectedBy: [], count: 995 },
    ] };
}

describe("severity model", () => {
  test("grades every instance: broken, in components, in designs, info only", () => {
    const ui = load(); const m = ui.auditModel(payload());
    assert.equal(m.broken, 1); assert.equal(m.inComp, 1); assert.equal(m.inDesign, 1); assert.equal(m.info, 2, "nested + graveyard-page are info only");
    assert.equal(m.linked, 995);
    assert.deepEqual(m.rows.map((r) => r.sev), ["broken", "component"], "sorted broken first, row severity is its worst instance");
  });
  test("graveyard-page instances never count as issues or appear in issues-by-page", () => {
    const ui = load(); const m = ui.auditModel(payload());
    assert.equal(m.byPage["Graveyard"], undefined);
  });
});

describe("rendering and export", () => {
  test("results render without error, with severity chips and the built-into note", () => {
    const ui = load(); ui.setAudit(payload()); ui.auditRender();
    const html = ui.store["au-results"].innerHTML;
    assert.ok(html.length > 1000, "panel has content"); assert.doesNotMatch(html, /Could not draw/);
    assert.match(html, /In components/); assert.match(html, /built into Card/);
    assert.equal(ui.store["badge-audit"].textContent, 3);
  });
  test("CSV has the severity columns and one row per flagged instance", () => {
    const ui = load(); ui.setAudit(payload());
    const lines = ui.auditCsv().split("\n");
    assert.match(lines[0], /^severity,component,builtInto,/);
    assert.equal(lines.length - 1, 5, "4 tooltip instances + 1 slot");
  });
  test("the canvas report payload mirrors the panel", () => {
    const ui = load(); ui.setAudit(payload()); const r = ui.auditReportPayload();
    // objects come from the vm sandbox (another realm), so compare by value, not by prototype
    assert.deepEqual(JSON.parse(JSON.stringify(r.counts)), { linked: 995, broken: 1, inComp: 1, inDesign: 1, info: 2 });
    assert.equal(r.issueCount, 3); assert.equal(r.rows[0].sev, "broken");
    assert.equal(r.rows[1].builtInto, "Card", "the report names the component the row is baked into, whichever instance carries it");
  });
});
