'use strict';
const ExcelJS = require('exceljs');
const { Resolver } = require('./resolver');
const { copyCellValue, detachSharedFormulaFollowers, runLogValue } = require('./utils');

const clone = (o) => (o === undefined ? undefined : JSON.parse(JSON.stringify(o)));
const GREEN = 'FF92D050', RED = 'FFFF0000';

function copySheet(src, dst) {
  // keep template widths; if the template has none, use readable defaults
  const fallback = [24, 24, 40, 22, 30, 22, 42];
  fallback.forEach((w, i) => { const sw = src.getColumn(i + 1).width; dst.getColumn(i + 1).width = sw || w; });
  src.eachRow({ includeEmpty: true }, (row, r) => {
    const drow = dst.getRow(r);
    if (row.height) drow.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, c) => {
      const d = drow.getCell(c);
      if (!cell.isMerged || cell.master === cell) d.value = copyCellValue(cell);
      d.style = clone(cell.style);
    });
  });
  // G-column merges belong to the template date's inactive inverters -> rebuilt per date below
  (src.model.merges || []).filter((m) => !/^G\d+:G\d+$/.test(m)).forEach((m) => dst.mergeCells(m));
  dst.views = clone(src.views);
}

function addMissingInverterSections(ws, data, rules) {
  const cfg = rules.inverter;
  const sectionRe = new RegExp(rules.inverter.sectionPattern);
  const normalizeInv = (value) => String(value).replace(/[^A-Z0-9]/gi, '').toUpperCase();
  const headers = data.rows(cfg.sheet)[0].slice(require('./utils').colToIndex(cfg.firstInverterCol));
  const dgrInverters = headers
    .map((header) => normalizeInv(header))
    .filter((header) => /^INV\d+$/.test(header));
  const existing = new Set();
  let firstBlockStart = 0;
  let firstBlockEnd = 0;

  for (let rowNumber = 1; rowNumber <= ws.rowCount; rowNumber++) {
    const cell = ws.getCell(`A${rowNumber}`);
    if (cell.isMerged && cell.master.address !== cell.address) continue;
    const value = cell.value;
    const text = value && typeof value === 'object' && value.richText
      ? value.richText.map((part) => part.text).join('')
      : value;
    const match = text && sectionRe.exec(String(text).trim());
    if (match) {
      existing.add(normalizeInv(match[1]));
      if (!firstBlockStart) firstBlockStart = rowNumber;
      else if (!firstBlockEnd) firstBlockEnd = rowNumber - 1;
    }
  }

  if (!firstBlockStart) return;
  if (!firstBlockEnd) firstBlockEnd = ws.rowCount;
  const templateRows = [];
  for (let rowNumber = firstBlockStart; rowNumber <= firstBlockEnd; rowNumber++) {
    templateRows.push(ws.getRow(rowNumber));
  }

  for (const inv of dgrInverters) {
    if (existing.has(inv)) continue;
    for (let i = 0; i < templateRows.length; i++) {
      const sourceRow = templateRows[i];
      const targetRow = ws.addRow([]);
      if (sourceRow.height) targetRow.height = sourceRow.height;
      for (let col = 1; col <= 7; col++) {
        const sourceCell = sourceRow.getCell(col);
        const targetCell = targetRow.getCell(col);
        targetCell.style = clone(sourceCell.style);
        if (col === 1) targetCell.value = i === 0 ? `Inverter Reading (${inv})` : null;
        else if (col === 2 || col === 4 || col === 6) targetCell.value = sourceCell.value;
      }
    }
  }
}

/**
 * Builds one validation sheet for (site, date) inside `outWb`.
 * Returns log rows for the run-log.
 */
function buildSheet(outWb, tplSheet, data, rules, { iso, site, siteCfg }) {
  const name = `${site || 'SITE'}_${iso}`.slice(0, 31);
  const ws = outWb.addWorksheet(name);
  copySheet(tplSheet, ws);
  addMissingInverterSections(ws, data, rules);

  const [y, m, d] = iso.split('-').map(Number);
  ws.getCell('D1').value = new Date(Date.UTC(y, m - 1, d));
  if (site) {
    for (const a of ['C2', 'E2']) {
      const v = String(ws.getCell(a).value || '');
      ws.getCell(a).value = v.replace(/_[^_()]+_\(PASS, FAIL\)/i, `_${site}_(PASS, FAIL)`);
    }
  }
  ws.getCell('F2').value = 'Expected';
  const expectedWidth = ws.getColumn(7).width || 42;
  const commentsWidth = ws.getColumn(6).width || 22;
  ws.getColumn(6).width = expectedWidth;
  ws.getColumn(7).width = commentsWidth;

  const resolver = new Resolver(data, rules, { iso, plantName: siteCfg && siteCfg.plantName });
  const secRe = new RegExp(rules.inverter.sectionPattern);
  const log = [];
  const last = ws.rowCount;

  // Reset tester columns (C, E) and expected column (F): keep formatting, drop old results.
  for (let r = 3; r <= last; r++) {
    for (const col of ['C', 'E']) { const c = ws.getCell(`${col}${r}`); c.value = null; c.fill = { type: 'pattern', pattern: 'none' }; }
    ws.getCell(`F${r}`).value = null;
    ws.getCell(`G${r}`).value = null;
  }

  // Pass / Fail dropdown + auto colouring so the tester's entries look like the template.
  for (const col of ['C', 'E']) {
    const range = `${col}3:${col}${last}`;
    ws.addConditionalFormatting({ ref: range, rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['"Pass"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: GREEN } } } },
      { type: 'cellIs', operator: 'equal', formulae: ['"Fail"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: RED } } } },
    ] });
    for (let r = 3; r <= last; r++) ws.getCell(`${col}${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Pass,Fail"'] };
  }

  // Walk rows, carrying Sub Module (A) forward.
  let subModule = '', inv = null, blockStart = 0;
  const blocks = [];
  for (let r = 3; r <= last; r++) {
    const aCell = ws.getCell(`A${r}`);
    // merged slave cells echo the master's value -> only the anchor row starts a new section
    const a = aCell.isMerged && aCell.master.address !== aCell.address ? null : aCell.value;
    const aText = a && typeof a === 'object' && a.richText ? a.richText.map((t) => t.text).join('') : a;
    const field = String(ws.getCell(`B${r}`).value ?? '').trim();
    if (aText) {
      subModule = String(aText).trim();
      const sm = secRe.exec(subModule);
      inv = sm ? sm[1] : null;
      if (inv) blocks.push({ inv, start: r, end: r });
    } else if (!field) { inv = null; continue; }
    if (!field) continue;

    if (inv) {
      blocks[blocks.length - 1].end = r;
      const res = resolver.resolveInverter(inv, field);
      log.push({ sheet: name, row: r, section: subModule, field, ...res });
      if (res.ok) {
        ws.getCell(`F${r}`).value = res.text;
        ws.getCell(`C${r}`).value = 'Pass';
        ws.getCell(`C${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
      }
      continue;
    }
    const comment = ws.getCell(`D${r}`).value || ws.getCell(`F${r}`).value;
    const res = resolver.resolveField(subModule, field, comment);
    log.push({ sheet: name, row: r, section: subModule, field, comment, ...res });
    if (res.ok) {
      ws.getCell(`F${r}`).value = res.text;
      ws.getCell(`C${r}`).value = 'Pass';
      ws.getCell(`C${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
    }
  }

  // Inactive inverters: one merged "Inverter Inactive" cell over the block (as in the template).
  for (const b of blocks) {
    if (!resolver.isInverterInactive(b.inv)) continue;
    for (let r = b.start; r <= b.end; r++) ws.getCell(`F${r}`).value = null;
    if (b.end > b.start) ws.mergeCells(`F${b.start}:F${b.end}`);
    ws.getCell(`F${b.start}`).value = rules.inverter.inactiveText;
    ws.getCell(`F${b.start}`).alignment = { horizontal: 'center', vertical: 'middle' };
    log.forEach((l) => {
      if (l.sheet === name && l.row >= b.start && l.row <= b.end && !l.skip) {
        l.ok = true;
        l.text = rules.inverter.inactiveText;
        l.note = 'INV_GEN = 0';
        ws.getCell(`C${l.row}`).value = 'Pass';
        ws.getCell(`C${l.row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
      }
    });
  }
  return { name, log };
}

function addRunLog(outWb, logs) {
  const ws = outWb.addWorksheet('Run_Log');
  ws.columns = [
    { header: 'Sheet', key: 'sheet', width: 20 }, { header: 'Row', key: 'row', width: 6 },
    { header: 'Sub Module', key: 'section', width: 28 }, { header: 'Field', key: 'field', width: 24 },
    { header: 'Comment (source hint)', key: 'comment', width: 28 }, { header: 'Resolved source', key: 'ref', width: 34 },
    { header: 'Raw value', key: 'raw', width: 18 }, { header: 'Written to F', key: 'text', width: 34 },
    { header: 'Status', key: 'status', width: 10 }, { header: 'Note', key: 'note', width: 44 },
  ];
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  for (const l of logs) {
    const status = l.ok ? 'OK' : l.skip ? 'SKIPPED' : 'ERROR';
    const row = ws.addRow({
      ...l,
      raw: runLogValue(l.raw),
      comment: l.comment && l.comment.richText ? '' : l.comment,
      status,
    });
    if (status === 'ERROR') row.getCell('status').font = { color: { argb: RED }, bold: true };
  }
  return ws;
}

async function generate({ template, data, rules, dates, site, siteCfg: providedSiteCfg, outFile }) {
  const tplWb = new ExcelJS.Workbook();
  await tplWb.xlsx.readFile(template);
  const tpl = tplWb.worksheets[0];
  const outWb = new ExcelJS.Workbook();
  const siteCfg = providedSiteCfg || rules.sites[site];
  const logs = [];
  for (const iso of dates) {
    const { log } = buildSheet(outWb, tpl, data, rules, { iso, site, siteCfg });
    logs.push(...log);
  }
  addRunLog(outWb, logs);
  detachSharedFormulaFollowers(outWb);
  await outWb.xlsx.writeFile(outFile);
  return logs;
}

module.exports = { generate };
