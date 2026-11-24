/**
 * Office Script: Add Additive Application
 *
 * Runs on a person's own workbook (for example "3_Workstream_Template_Miko.xlsx").
 * - Adds a new tab for the application, copied from the blank "Additive Application"
 *   form in A1:B14 of the "Form Template" sheet, or of "Workstream Example" if there's no template sheet.
 * - Fills in every row and puts the part photo in the Image row.
 * - Adds (or updates) a line on the "Summary" tab, with a link to the new tab.
 * Posting the same entry again from the app updates its tab instead of making a duplicate.
 *
 * Called by the Teams workflow. Returns the name of the tab it wrote.
 */

interface Rows {
  originator: string; department: string; workstreamOwner: string; appName: string; appDesc: string;
  issue: string; partInfo: string; additiveNotes: string; customerUsage: string; urgency: string;
  confidenceText: string; customerInput: string;
}
interface Detail { n: number; label: string; value: string; }
interface Photo { base64: string; w: number; h: number; }
interface Entry {
  v: number; id: string; createdAt: string; updatedAt: string; originator: string; notesFolder: string;
  workstreamOwner: string; rows: Rows; urgency: number; confidence: number; eau: string; suggestedAM: string;
  details: Detail[]; otherNotes: string; photo: Photo | null;
}

const SUMMARY = "Summary";
const HEADERS = ["#", "Captured", "Originator", "Department/Area/Product Line", "Workstream Owner", "Application Name",
  "Urgency (1-10)", "Confidence (1-10)", "EAU", "Current Issue/Challenge", "Form", "Entry ID", "Tab"];
const ID_COL = 11, TAB_COL = 12; // zero-based columns L and M (kept hidden)

function main(workbook: ExcelScript.Workbook, entryJson: string): string {
  const e = JSON.parse(entryJson) as Entry;
  if (!e || !e.id) { throw new Error("No entry data was received."); }

  const template = findTemplate(workbook);
  const summary = getSummary(workbook);

  // Was this entry posted before?
  const existingRow = findRow(summary, e.id);
  let tabName: string;
  let position = -1;
  if (existingRow > 0) {
    tabName = String(summary.getCell(existingRow, TAB_COL).getValue());
    const old = workbook.getWorksheet(tabName);
    if (old) { position = old.getPosition(); old.delete(); }
  } else {
    const count = dataRowCount(summary);
    tabName = uniqueSheetName(workbook, pad(count + 1) + " " + (e.rows.appName || "Untitled"));
  }

  const ws = workbook.addWorksheet(tabName);
  if (position >= 0) { ws.setPosition(position); }
  fillForm(ws, template, e);

  writeSummaryRow(summary, existingRow, e, tabName);
  summary.setPosition(0);
  return tabName;
}

/* ---------------- form tab ---------------- */
function findTemplate(workbook: ExcelScript.Workbook): ExcelScript.Range {
  const sheet = workbook.getWorksheet("Form Template") || workbook.getWorksheet("Workstream Example") || workbook.getWorksheets()[0];
  return sheet.getRange("A1:B14");
}

function fillForm(ws: ExcelScript.Worksheet, template: ExcelScript.Range, e: Entry) {
  const tplSheet = template.getWorksheet();
  ws.getRange("A1:B14").copyFrom(template, ExcelScript.RangeCopyType.all, false, false);
  ws.getRange("A:A").getFormat().setColumnWidth(tplSheet.getRange("A:A").getFormat().getColumnWidth());
  ws.getRange("B:B").getFormat().setColumnWidth(tplSheet.getRange("B:B").getFormat().getColumnWidth());
  ws.getRange("C:C").getFormat().setColumnWidth(70);
  ws.getRange("A1").getFormat().setRowHeight(tplSheet.getRange("A1").getFormat().getRowHeight());

  const r = e.rows;
  const values: (string | number)[] = [r.originator, r.department, r.workstreamOwner, r.appName, r.appDesc, r.issue,
    r.partInfo, r.additiveNotes, r.customerUsage, e.urgency || "", r.confidenceText, r.customerInput, ""];
  for (let i = 0; i < values.length; i++) {
    ws.getCell(i + 1, 1).setValue(values[i]);
  }
  const body = ws.getRange("B2:B13");
  body.getFormat().setWrapText(true);
  body.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  ws.getRange("A2:B13").getFormat().autofitRows();

  // Urgency cell: green (1) to red (10)
  const urg = ws.getRange("B11");
  if (e.urgency) {
    urg.getFormat().getFill().setColor(urgencyColor(e.urgency));
    urg.getFormat().getFont().setBold(true);
    urg.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  }

  // Part photo in the Image row (B14)
  const imgCell = ws.getRange("B14");
  if (e.photo && e.photo.base64) {
    const colW = ws.getRange("B:B").getFormat().getColumnWidth();
    let w = Math.min(colW - 10, 360);
    let h = w * (e.photo.h / Math.max(1, e.photo.w));
    if (h > 300) { w = w * 300 / h; h = 300; }
    imgCell.getFormat().setRowHeight(Math.min(409, h + 10));
    const shape = ws.addImage(e.photo.base64);
    shape.setLockAspectRatio(true);
    shape.setWidth(w);
    shape.setHeight(h);
    shape.setLeft(imgCell.getLeft() + 5);
    shape.setTop(imgCell.getTop() + 5);
    shape.setName("Part photo");
  } else {
    imgCell.setValue("(No part photo)");
  }

  // Link back to the Summary tab
  const back = ws.getRange("C1");
  back.setHyperlink({ documentReference: "'" + SUMMARY + "'!A1", textToDisplay: "← Summary" });

  // Notebook details, exactly as captured
  const labelFill = tplSheet.getRange("A2").getFormat().getFill().getColor();
  const lines: string[][] = [];
  lines.push(["Captured", formatDate(e.createdAt)]);
  lines.push(["Last edited", formatDate(e.updatedAt)]);
  for (const d of e.details) { lines.push([d.n + ". " + d.label, d.value || ""]); }
  if (e.otherNotes) { lines.push(["Other notes", e.otherNotes]); }

  const start = 16; // row 16, leaving row 15 blank like the template
  const head = ws.getRange("A" + start + ":B" + start);
  head.merge(false);
  head.setValue("Notebook Details (as captured)");
  head.getFormat().getFont().setBold(true);
  head.getFormat().getFill().setColor(labelFill);
  const block = ws.getRange("A" + (start + 1) + ":B" + (start + lines.length));
  block.setValues(lines);
  block.getFormat().setWrapText(true);
  block.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  ws.getRange("A" + (start + 1) + ":A" + (start + lines.length)).getFormat().getFill().setColor(labelFill);
  borders(ws.getRange("A" + start + ":B" + (start + lines.length)));
  block.getFormat().autofitRows();
}

/* ---------------- summary tab ---------------- */
function getSummary(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  let s = workbook.getWorksheet(SUMMARY);
  if (!s) {
    s = workbook.addWorksheet(SUMMARY);
    const header = s.getRangeByIndexes(0, 0, 1, HEADERS.length);
    header.setValues([HEADERS]);
    header.getFormat().getFont().setBold(true);
    header.getFormat().getFont().setColor("#FFFFFF");
    header.getFormat().getFill().setColor("#282156");
    header.getFormat().setWrapText(true);
    header.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    const widths = [28, 95, 95, 130, 105, 170, 62, 70, 60, 280, 45, 60, 60];
    for (let i = 0; i < widths.length; i++) {
      s.getRangeByIndexes(0, i, 1, 1).getFormat().setColumnWidth(widths[i]);
    }
    s.getRange("L:M").setColumnHidden(true);
    s.getFreezePanes().freezeRows(1);
  }
  return s;
}

function findRow(summary: ExcelScript.Worksheet, id: string): number {
  const used = summary.getUsedRange(true);
  if (!used) { return -1; }
  const vals = used.getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i].length > ID_COL && String(vals[i][ID_COL]) === id) { return i; }
  }
  return -1;
}

function dataRowCount(summary: ExcelScript.Worksheet): number {
  const used = summary.getUsedRange(true);
  if (!used) { return 0; }
  return Math.max(0, used.getRowCount() - 1);
}

function writeSummaryRow(summary: ExcelScript.Worksheet, existingRow: number, e: Entry, tabName: string) {
  const row = existingRow > 0 ? existingRow : dataRowCount(summary) + 1;
  const num = existingRow > 0 ? summary.getCell(row, 0).getValue() : row;
  const values: (string | number | boolean)[] = [num, formatDate(e.createdAt), e.rows.originator, e.rows.department,
    e.rows.workstreamOwner, e.rows.appName, e.urgency || "", e.confidence || "", e.eau, e.rows.issue, "Open", e.id, tabName];
  const range = summary.getRangeByIndexes(row, 0, 1, values.length);
  range.setValues([values]);
  range.getFormat().setWrapText(true);
  range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  borders(range);
  summary.getCell(row, 10).setHyperlink({ documentReference: "'" + tabName.replace(/'/g, "''") + "'!A1", textToDisplay: "Open" });
  const urg = summary.getCell(row, 6);
  urg.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  if (e.urgency) { urg.getFormat().getFill().setColor(urgencyColor(e.urgency)); }
  summary.getCell(row, 7).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  range.getFormat().autofitRows();
}

/* ---------------- helpers ---------------- */
function borders(range: ExcelScript.Range) {
  const edges = [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom, ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight];
  if (range.getRowCount() > 1) { edges.push(ExcelScript.BorderIndex.insideHorizontal); }
  if (range.getColumnCount() > 1) { edges.push(ExcelScript.BorderIndex.insideVertical); }
  for (const edge of edges) {
    const b = range.getFormat().getRangeBorder(edge);
    b.setStyle(ExcelScript.BorderLineStyle.continuous);
    b.setWeight(ExcelScript.BorderWeight.thin);
    b.setColor("#000000");
  }
}

function uniqueSheetName(workbook: ExcelScript.Workbook, wanted: string): string {
  let base = wanted.replace(/[\[\]:*?\/\\]/g, " ").replace(/\s+/g, " ").trim().replace(/^'+|'+$/g, "");
  if (!base) { base = "Application"; }
  base = base.substring(0, 31).trim();
  const taken = workbook.getWorksheets().map(w => w.getName().toLowerCase());
  let name = base;
  let n = 2;
  while (taken.indexOf(name.toLowerCase()) >= 0 || name.toLowerCase() === SUMMARY.toLowerCase()) {
    const suffix = " (" + n + ")";
    name = base.substring(0, 31 - suffix.length).trim() + suffix;
    n++;
  }
  return name;
}

function pad(n: number): string { return n < 10 ? "0" + n : String(n); }

function formatDate(iso: string): string {
  if (!iso) { return ""; }
  const d = new Date(iso);
  if (isNaN(d.getTime())) { return iso; }
  // Shown in US Eastern time, where the team works.
  try {
    return d.toLocaleString("en-US", { timeZone: "America/New_York", year: "numeric", month: "2-digit", day: "2-digit", hour: "numeric", minute: "2-digit" });
  } catch (err) {
    return d.toISOString().substring(0, 16).replace("T", " ") + " UTC";
  }
}

function urgencyColor(n: number): string {
  const v = Math.max(1, Math.min(10, Math.round(n)));
  const h = 120 * (10 - v) / 9, s = 0.7, l = 0.68;
  const k = (m: number) => (m + h / 30) % 12;
  const a = s * Math.min(l, 1 - l);
  const f = (m: number) => l - a * Math.max(-1, Math.min(k(m) - 3, Math.min(9 - k(m), 1)));
  const hex = (x: number) => { const t = Math.round(x * 255).toString(16); return t.length < 2 ? "0" + t : t; };
  return "#" + hex(f(0)) + hex(f(8)) + hex(f(4));
}
