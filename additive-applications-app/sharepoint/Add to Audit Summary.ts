/**
 * Office Script: Add to Audit Summary
 *
 * Runs on the shared "Audit Summary.xlsx" in the current audit folder.
 * Adds one row per application from everyone on the team, with a link to the person's workbook.
 * Posting the same entry again from the app updates its row instead of adding a duplicate.
 *
 * Called by the Teams workflow after "Add Additive Application".
 *   entryJson    – the entry data the app sends
 *   tabName      – the tab the first script created in the person's workbook
 *   workbookLink – web link to the person's workbook (optional)
 */

interface Rows {
  originator: string; department: string; workstreamOwner: string; appName: string; appDesc: string;
  issue: string; partInfo: string; additiveNotes: string; customerUsage: string; urgency: string;
  confidenceText: string; customerInput: string;
}
interface Entry {
  v: number; id: string; createdAt: string; updatedAt: string; originator: string; notesFolder: string;
  workstreamOwner: string; rows: Rows; urgency: number; confidence: number; eau: string; suggestedAM: string;
  otherNotes: string;
}

const SHEET = "Audit Summary";
const HEADERS = ["#", "Captured", "Originator", "Notes Folder", "Department/Area/Product Line", "Workstream Owner",
  "Application Name", "Urgency (1-10)", "Confidence (1-10)", "EAU", "Current Issue/Challenge", "Suggested AM Material/Process",
  "Workbook", "Tab", "Entry ID"];
const LINK_COL = 12, ID_COL = 14;

function main(workbook: ExcelScript.Workbook, entryJson: string, tabName: string, workbookLink: string): string {
  const e = JSON.parse(entryJson) as Entry;
  if (!e || !e.id) { throw new Error("No entry data was received."); }
  const s = getSheet(workbook);

  let row = findRow(s, e.id);
  const isNew = row < 0;
  if (isNew) { row = dataRowCount(s) + 1; }
  const num = isNew ? row : s.getCell(row, 0).getValue();

  const values: (string | number | boolean)[] = [num, formatDate(e.createdAt), e.rows.originator, e.notesFolder,
    e.rows.department, e.rows.workstreamOwner, e.rows.appName, e.urgency || "", e.confidence || "", e.eau,
    e.rows.issue, e.suggestedAM, (e.notesFolder || "Person") + " workbook", tabName || "", e.id];
  const range = s.getRangeByIndexes(row, 0, 1, values.length);
  range.setValues([values]);
  range.getFormat().setWrapText(true);
  range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  borders(range);

  if (workbookLink) {
    s.getCell(row, LINK_COL).setHyperlink({ address: workbookLink, textToDisplay: (e.notesFolder || "Person") + " workbook" });
  }
  const urg = s.getCell(row, 7);
  urg.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  if (e.urgency) { urg.getFormat().getFill().setColor(urgencyColor(e.urgency)); }
  s.getCell(row, 8).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  range.getFormat().autofitRows();
  return isNew ? "added" : "updated";
}

function getSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  let s = workbook.getWorksheet(SHEET);
  if (!s) {
    const first = workbook.getWorksheets()[0];
    const empty = first && workbook.getWorksheets().length === 1 && isBlank(first);
    if (empty) { s = first; s.setName(SHEET); } else { s = workbook.addWorksheet(SHEET); }
    const header = s.getRangeByIndexes(0, 0, 1, HEADERS.length);
    header.setValues([HEADERS]);
    header.getFormat().getFont().setBold(true);
    header.getFormat().getFont().setColor("#FFFFFF");
    header.getFormat().getFill().setColor("#282156");
    header.getFormat().setWrapText(true);
    header.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    const widths = [28, 95, 95, 70, 130, 105, 170, 62, 70, 60, 280, 150, 95, 110, 60];
    for (let i = 0; i < widths.length; i++) {
      s.getRangeByIndexes(0, i, 1, 1).getFormat().setColumnWidth(widths[i]);
    }
    s.getRangeByIndexes(0, ID_COL, 1, 1).getEntireColumn().setColumnHidden(true);
    s.getFreezePanes().freezeRows(1);
    s.setPosition(0);
  }
  return s;
}

function isBlank(sheet: ExcelScript.Worksheet): boolean {
  const used = sheet.getUsedRange(true);
  if (!used) { return true; }
  return used.getRowCount() === 1 && used.getColumnCount() === 1 && String(used.getValue()) === "";
}

function findRow(s: ExcelScript.Worksheet, id: string): number {
  const used = s.getUsedRange(true);
  if (!used) { return -1; }
  const vals = used.getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i].length > ID_COL && String(vals[i][ID_COL]) === id) { return i; }
  }
  return -1;
}

function dataRowCount(s: ExcelScript.Worksheet): number {
  const used = s.getUsedRange(true);
  if (!used) { return 0; }
  return Math.max(0, used.getRowCount() - 1);
}

function borders(range: ExcelScript.Range) {
  const edges = [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom, ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight, ExcelScript.BorderIndex.insideVertical];
  for (const edge of edges) {
    const b = range.getFormat().getRangeBorder(edge);
    b.setStyle(ExcelScript.BorderLineStyle.continuous);
    b.setWeight(ExcelScript.BorderWeight.thin);
    b.setColor("#000000");
  }
}

function formatDate(iso: string): string {
  if (!iso) { return ""; }
  const d = new Date(iso);
  if (isNaN(d.getTime())) { return iso; }
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
