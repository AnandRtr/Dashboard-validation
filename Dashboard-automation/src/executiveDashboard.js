'use strict';

const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const {
  colToIndex,
  isNum,
  formatSourceValue,
  serialToIso,
  copyCellValue,
  detachSharedFormulaFollowers,
  runLogValue,
} = require('./utils');

const RED = 'FFFF0000';
const EXECUTIVE_SHEET_DATE_COLUMNS = {
  BRIDGE_GAP_ANALYSIS: 'A',
  DIGITAL_TWIN_ANALYSIS: 'O',
};

function cellText(value) {
  if (value && typeof value === 'object' && value.richText) {
    return value.richText.map((part) => part.text).join('');
  }
  if (value && typeof value === 'object' && Object.prototype.hasOwnProperty.call(value, 'result')) {
    return cellText(value.result);
  }
  return value === null || value === undefined ? '' : String(value);
}

function findDateColumn(rows, sheet, rules, iso) {
  const lookupSheet = String(sheet || '').trim();
  const aliases = (rules && rules.sheetAliases ? Object.values(rules.sheetAliases) : []);
  const alias = aliases.find((item) => String(item.sheet || '').toLowerCase() === lookupSheet.toLowerCase());
  if (alias && alias.dateCol) return colToIndex(alias.dateCol);

  const configuredDateColumn = Object.entries(EXECUTIVE_SHEET_DATE_COLUMNS)
    .find(([name]) => name.toLowerCase() === lookupSheet.toLowerCase());
  if (configuredDateColumn) return colToIndex(configuredDateColumn[1]);

  for (const row of rows.slice(0, 25)) {
    const dateHeader = row.findIndex((value) => /\bdate\b/i.test(String(value || '')));
    if (dateHeader >= 0) return dateHeader;
  }

  const dateColumns = new Set();
  for (const row of rows) {
    row.forEach((value, index) => {
      if (dateToIso(value) === iso) dateColumns.add(index);
    });
  }
  if (dateColumns.size === 1) return [...dateColumns][0];
  return -1;
}

function dateToIso(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'string') {
    const text = value.trim();
    if (!text) return null;
    const asNumber = Number(text);
    if (!Number.isNaN(asNumber) && Number.isFinite(asNumber) && asNumber >= 20000 && asNumber <= 70000) {
      return serialToIso(asNumber);
    }
  }
  if (isNum(value) && value >= 20000 && value <= 70000) {
    return serialToIso(value);
  }
  if (typeof value !== 'string') return null;

  const text = value.trim();
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$/.exec(text);
  if (match) {
    const [, year, month, day] = match;
    const iso = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
    const parsed = new Date(`${iso}T00:00:00.000Z`);
    return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === iso ? iso : null;
  }
  match = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:[ T].*)?$/.exec(text);
  if (!match) return null;
  const [, day, month, year] = match;
  const iso = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
  const parsed = new Date(`${iso}T00:00:00.000Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === iso ? iso : null;
}

function readDgrWorkbook(file, rules) {
  const workbook = XLSX.readFile(file, { cellDates: false, cellFormula: false, cellNF: true, cellStyles: false });
  const sheets = new Map();
  for (const name of workbook.SheetNames) {
    const worksheet = workbook.Sheets[name];
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const numberFormats = new Map();
    for (const address of Object.keys(worksheet)) {
      if (address.startsWith('!')) continue;
      const cell = worksheet[address];
      if (cell && cell.z) {
        const position = XLSX.utils.decode_cell(address);
        numberFormats.set(`${position.r - range.s.r}:${position.c - range.s.c}`, cell.z);
      }
    }
    sheets.set(name.toLowerCase(), {
      name,
      rows: XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: true,
        defval: null,
        blankrows: true,
      }),
      numberFormats,
    });
  }
  return { sheets, rules };
}

function getConfigEntry(config, lookupKey, fallbackName) {
  if (!config) return fallbackName;
  const key = String(lookupKey || '').trim();
  if (config[key]) return config[key];
  const match = Object.entries(config).find(([name]) => name.toLowerCase() === key.toLowerCase());
  return match ? match[1] : fallbackName;
}

function normalizeLabel(value) {
  return cellText(value).trim().replace(/\s+/g, ' ').toLowerCase();
}

function findReferenceMapping(sheet, section, field) {
  let currentSection = '';
  const matches = [];
  for (let rowNumber = 3; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const sectionCell = row.getCell(2);
    if (!(sectionCell.isMerged && sectionCell.master.address !== sectionCell.address) && sectionCell.value) {
      currentSection = normalizeLabel(sectionCell.value);
    }
    if (currentSection === normalizeLabel(section)
      && normalizeLabel(row.getCell(3).value) === normalizeLabel(field)) {
      matches.push({ rowNumber, mapping: cellText(row.getCell(5).value).trim() });
    }
  }
  if (matches.length > 1) {
    return { found: true, ok: false, note: `Multiple reference mappings match ${section} / ${field}.` };
  }
  return matches.length ? { found: true, ...matches[0] } : { found: false };
}

function comparisonCellAddresses(mapping, section) {
  const isChart = normalizeLabel(section) === 'chart';
  const addresses = isChart
    ? [...String(mapping || '').matchAll(/[^:;]+:\s*(\$?[A-Z]{1,3}\$?\d+)/gi)].map((match) => match[1])
    : [String(mapping || '').trim()];
  return addresses
    .map((address) => /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(address))
    .filter(Boolean)
    .map((match) => `${match[1].toUpperCase()}${match[2]}`);
}

function comparisonMappings(sheet) {
  const mappings = [];
  let currentSection = '';
  for (let rowNumber = 3; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const sectionCell = row.getCell(2);
    if (!(sectionCell.isMerged && sectionCell.master.address !== sectionCell.address) && sectionCell.value) {
      currentSection = normalizeLabel(sectionCell.value);
    }
    if (!['chart', 'kpi comparision', 'kpi comparison'].includes(currentSection)) continue;
    mappings.push({
      section: currentSection,
      mapping: cellText(row.getCell(5).value).trim(),
    });
  }
  return mappings;
}

function mappedCellCoverage(sheet, mappings) {
  const addresses = new Set(mappings.flatMap(({ mapping, section }) => comparisonCellAddresses(mapping, section)));
  let populated = 0;
  for (const address of addresses) {
    const value = sheet.getCell(address).value;
    const resolved = value && typeof value === 'object' && Object.prototype.hasOwnProperty.call(value, 'result')
      ? value.result
      : value;
    if (resolved !== null && resolved !== undefined && resolved !== '') populated++;
  }
  return { populated, total: addresses.size };
}

function selectComparisonSheet(workbook, referenceSheet, iso) {
  const mappings = comparisonMappings(referenceSheet);
  const rankedSheets = workbook.worksheets
    .map((sheet, index) => {
      const coverage = mappedCellCoverage(sheet, mappings);
      return {
        sheet,
        index,
        ...coverage,
        dateMatch: reportDateFromSheetName(sheet.name) === iso,
        namedForComparison: /kpi|comparison|chart/i.test(sheet.name),
      };
    })
    .sort((a, b) => b.populated - a.populated
      || Number(b.dateMatch) - Number(a.dateMatch)
      || Number(b.namedForComparison) - Number(a.namedForComparison)
      || a.index - b.index);
  return rankedSheets[0].sheet;
}

function formatExpectedValue(value, numberFormat) {
  const result = value && typeof value === 'object' && Object.prototype.hasOwnProperty.call(value, 'result')
    ? value.result
    : value;
  if (result === null || result === undefined || result === '') return null;
  return isNum(result) ? formatSourceValue(result, numberFormat) : cellText(result);
}

function evaluateFormula(sheet, formula, visiting) {
  const expression = String(formula || '').replace(/^=/, '');
  const tokenPattern = /\s*(?:(\d+(?:\.\d*)?(?:[eE][+-]?\d+)?)|(\$?[A-Za-z]{1,3}\$?\d+)|([()+\-*/^]))/gy;
  const tokens = [];
  let offset = 0;
  while (offset < expression.length) {
    tokenPattern.lastIndex = offset;
    const match = tokenPattern.exec(expression);
    if (!match) throw new Error(`Unsupported formula syntax near "${expression.slice(offset)}".`);
    tokens.push(match[1] ? Number(match[1]) : match[2] ? match[2] : match[3]);
    offset = tokenPattern.lastIndex;
  }

  let position = 0;
  const parsePrimary = () => {
    const token = tokens[position++];
    if (typeof token === 'number') return token;
    if (token === '(') {
      const value = parseExpression();
      if (tokens[position++] !== ')') throw new Error('Missing closing parenthesis.');
      return value;
    }
    if (typeof token === 'string' && /^\$?[A-Za-z]{1,3}\$?\d+$/.test(token)) {
      const address = token.replace(/\$/g, '').toUpperCase();
      const resolved = resolveComparisonCellValue(sheet, address, visiting, true);
      if (!resolved.ok) throw new Error(resolved.note);
      const numeric = resolved.value === '' || resolved.value === null ? 0 : Number(resolved.value);
      if (!Number.isFinite(numeric)) throw new Error(`Referenced cell ${address} is not numeric.`);
      return numeric;
    }
    throw new Error(`Unexpected formula token "${token ?? ''}".`);
  };
  const parsePower = () => {
    const left = parsePrimary();
    if (tokens[position] === '^') {
      position++;
      return left ** parseUnary();
    }
    return left;
  };
  const parseUnary = () => {
    if (tokens[position] === '+') {
      position++;
      return parseUnary();
    }
    if (tokens[position] === '-') {
      position++;
      return -parseUnary();
    }
    return parsePower();
  };
  const parseTerm = () => {
    let value = parseUnary();
    while (tokens[position] === '*' || tokens[position] === '/') {
      const operator = tokens[position++];
      const right = parseUnary();
      if (operator === '/' && right === 0) throw new Error('Formula divides by zero.');
      value = operator === '*' ? value * right : value / right;
    }
    return value;
  };
  function parseExpression() {
    let value = parseTerm();
    while (tokens[position] === '+' || tokens[position] === '-') {
      const operator = tokens[position++];
      const right = parseTerm();
      value = operator === '+' ? value + right : value - right;
    }
    return value;
  }

  if (!tokens.length) throw new Error('Formula is empty.');
  const value = parseExpression();
  if (position !== tokens.length) throw new Error(`Unexpected formula token "${tokens[position]}".`);
  if (!Number.isFinite(value)) throw new Error('Formula did not produce a finite number.');
  return value;
}

function resolveComparisonCellValue(sheet, address, visiting = new Set(), blankAsZero = false) {
  const cell = sheet.getCell(address);
  const value = cell.value;
  if (value === null || value === undefined || value === '') {
    return blankAsZero ? { ok: true, value: 0 } : { ok: false, note: 'Mapped comparison cell is empty.' };
  }
  if (value && typeof value === 'object') {
    if (Object.prototype.hasOwnProperty.call(value, 'result') && value.result !== undefined) {
      return { ok: true, value: value.result };
    }
    const formula = cell.formula || value.formula;
    if (formula) {
      const key = `${sheet.name}!${cell.address}`;
      if (visiting.has(key)) return { ok: false, note: `Formula cycle detected at ${key}.` };
      const nextVisiting = new Set(visiting);
      nextVisiting.add(key);
      try {
        return { ok: true, value: evaluateFormula(sheet, formula, nextVisiting) };
      } catch (error) {
        return { ok: false, note: `Cannot resolve formula in ${key}: ${error.message}` };
      }
    }
    if (value.sharedFormula) {
      return { ok: false, note: `Shared formula in ${sheet.name}!${cell.address} has no cached value or resolvable master.` };
    }
  }
  return { ok: true, value };
}

function getMappedCellValue(sheet, address) {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(String(address || '').trim());
  if (!match) return { ok: false, note: `Invalid comparison cell address "${address}".` };
  const normalizedAddress = `${match[1].toUpperCase()}${match[2]}`;
  const cell = sheet.getCell(normalizedAddress);
  const resolved = resolveComparisonCellValue(sheet, normalizedAddress);
  if (!resolved.ok) return { ok: false, ref: `${sheet.name}!${normalizedAddress}`, note: resolved.note };
  const formatted = formatExpectedValue(resolved.value, cell.numFmt);
  if (formatted === null) {
    return { ok: false, ref: `${sheet.name}!${normalizedAddress}`, note: 'Mapped comparison cell is empty.' };
  }
  return {
    ok: true,
    ref: `${sheet.name}!${normalizedAddress}`,
    raw: resolved.value,
    value: formatted,
  };
}

function resolveMappedReference(sheet, mapping, section) {
  if (!mapping) return { ok: false, skip: true, note: 'No comparison cell mapping in the reference sheet.' };
  const isChart = normalizeLabel(section) === 'chart';
  if (!isChart) {
    const result = getMappedCellValue(sheet, mapping);
    return result.ok
      ? { ...result, text: `Expected:${result.value}` }
      : result;
  }

  const entries = [...mapping.matchAll(/([^:;]+):\s*(\$?[A-Z]{1,3}\$?\d+)/gi)];
  if (!entries.length) return { ok: false, note: `No chart cell addresses found in reference mapping "${mapping}".` };
  const values = [];
  for (const [, label, address] of entries) {
    const result = getMappedCellValue(sheet, address);
    if (!result.ok) return { ...result, note: `${label.trim()}: ${result.note}` };
    values.push({ label: label.trim(), ...result });
  }
  return {
    ok: true,
    ref: values.map((entry) => entry.ref).join('; '),
    raw: values.map((entry) => `${entry.label}: ${entry.raw}`).join('\n'),
    text: `Expected:${values.map((entry) => `${entry.label}: ${entry.value}`).join('\n')}`,
  };
}

function reportDateFromSheetName(name) {
  const iso = /\b(\d{4}-\d{1,2}-\d{1,2})\b/.exec(name);
  if (iso) return dateToIso(iso[1]);
  const dmy = /\b(\d{1,2})[-/](\d{1,2})[-/](\d{4})\b/.exec(name);
  if (!dmy) return null;
  return dateToIso(`${dmy[1]}-${dmy[2]}-${dmy[3]}`);
}

function resolveSource(workbook, source, iso, plantName) {
  const match = /^\s*([A-Za-z][A-Za-z0-9_]*)_([A-Za-z]{1,3})(?![A-Za-z0-9])/.exec(String(source || ''));
  if (!match) return { ok: false, skip: true, note: 'No source sheet/column marker in the reference cell.' };

  const alias = workbook.rules.sheetAliases[match[1].toLowerCase()] || getConfigEntry(workbook.rules.sheetAliases, match[1], null);
  const executiveAliases = {
    bga: 'BRIDGE_GAP_ANALYSIS',
    dta: 'DIGITAL_TWIN_ANALYSIS',
  };
  const sheetName = alias ? alias.sheet : executiveAliases[match[1].toLowerCase()] || match[1];
  const sheet = workbook.sheets.get(sheetName.toLowerCase()) || [...workbook.sheets.values()].find((candidate) => candidate.name.toLowerCase() === String(sheetName).trim().toLowerCase());
  if (!sheet) return { ok: false, note: `Source sheet "${sheetName}" is missing from the DGR workbook.` };

  const dateIndex = findDateColumn(sheet.rows, sheet.name, workbook.rules, iso);
  if (dateIndex < 0) {
    return { ok: false, note: `Could not identify a date column in source sheet "${sheet.name}".` };
  }

  const plantColumn = getConfigEntry(workbook.rules.plantNameColumn, sheet.name, null);
  const plantIndex = plantColumn ? colToIndex(plantColumn) : -1;
  const rowIndex = sheet.rows.findIndex((candidate, index) => {
    if (index === 0 || dateToIso(candidate[dateIndex]) !== iso) return false;
    return !plantName || plantIndex < 0
      || String(candidate[plantIndex] || '').trim().toLowerCase() === plantName.toLowerCase();
  });
  if (rowIndex < 0) return { ok: false, ref: `${sheet.name}!${match[2].toUpperCase()}`, note: `No row for ${iso} in ${sheet.name}.` };

  const column = match[2].toUpperCase();
  const columnIndex = colToIndex(column);
  const raw = sheet.rows[rowIndex][columnIndex];
  const ref = `${sheet.name}!${column}`;
  if (raw === null || raw === undefined || raw === '') return { ok: false, ref, note: `${ref} is empty.` };

  const numberFormat = sheet.numberFormats.get(`${rowIndex}:${columnIndex}`);
  return { ok: true, ref, raw, text: `Expected:${formatSourceValue(raw, numberFormat)}` };
}

function copyTemplateSheet(source, destination) {
  source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const target = destination.getRow(rowNumber);
    if (row.height) target.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, column) => {
      const output = target.getCell(column);
      if (!cell.isMerged || cell.master === cell) output.value = copyCellValue(cell);
      output.style = JSON.parse(JSON.stringify(cell.style));
    });
  });
  for (const merge of source.model.merges || []) destination.mergeCells(merge);
  source.columns.forEach((column, index) => {
    if (column.width) destination.getColumn(index + 1).width = column.width;
  });
  destination.views = JSON.parse(JSON.stringify(source.views || []));
}

function addRunLog(workbook, logs) {
  const sheet = workbook.addWorksheet('Run_Log');
  sheet.columns = [
    { header: 'Sheet', key: 'sheet', width: 24 },
    { header: 'Row', key: 'row', width: 8 },
    { header: 'Sub Module', key: 'section', width: 32 },
    { header: 'Field', key: 'field', width: 36 },
    { header: 'Source marker', key: 'source', width: 24 },
    { header: 'Resolved source', key: 'ref', width: 28 },
    { header: 'Raw value', key: 'raw', width: 18 },
    { header: 'Written to G', key: 'text', width: 32 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Note', key: 'note', width: 52 },
  ];
  sheet.getRow(1).font = { bold: true };
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  for (const log of logs) {
    const row = sheet.addRow({
      ...log,
      raw: runLogValue(log.raw),
      status: log.ok ? 'OK' : log.skip ? 'SKIPPED' : 'ERROR',
    });
    if (!log.ok && !log.skip) row.getCell('status').font = { color: { argb: RED }, bold: true };
  }
}

async function generateExecutive({ dgrFile, template, comparison, rules, dates, site, siteCfg, outFile }) {
  const workbookData = readDgrWorkbook(dgrFile, rules);
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.readFile(template);
  const templateSheet = templateWorkbook.worksheets[0];
  if (!templateSheet) throw new Error(`No worksheet found in executive template: ${template}`);
  const comparisonWorkbook = new ExcelJS.Workbook();
  await comparisonWorkbook.xlsx.readFile(comparison);
  if (!comparisonWorkbook.worksheets.length) {
    throw new Error(`No worksheets found in KPI comparison workbook: ${comparison}`);
  }

  const outputWorkbook = new ExcelJS.Workbook();
  const logs = [];
  const plantName = (siteCfg && siteCfg.plantName) || (rules.sites[site] && rules.sites[site].plantName);

  for (const iso of dates) {
    const sheetName = `${site || 'SITE'}_${iso}`.slice(0, 31);
    const comparisonSheet = selectComparisonSheet(comparisonWorkbook, templateSheet, iso);
    const output = outputWorkbook.addWorksheet(sheetName);
    copyTemplateSheet(templateSheet, output);
    output.getCell('D1').value = new Date(`${iso}T00:00:00.000Z`);

    let subModule = '';
    for (let rowNumber = 3; rowNumber <= output.rowCount; rowNumber++) {
      const field = cellText(output.getCell(`C${rowNumber}`).value).trim();
      if (!field) continue;

      const sectionValue = output.getCell(`B${rowNumber}`);
      if (!(sectionValue.isMerged && sectionValue.master.address !== sectionValue.address) && sectionValue.value) {
        subModule = cellText(sectionValue.value).trim();
      }
      const source = cellText(output.getCell(`E${rowNumber}`).value).trim();
      const useComparison = ['kpi comparision', 'kpi comparison', 'chart'].includes(normalizeLabel(subModule));
      const referenceMapping = useComparison
        ? findReferenceMapping(templateSheet, subModule, field)
        : { found: false };
      if (source || referenceMapping.found || useComparison) output.getCell(`G${rowNumber}`).value = null;
      const result = useComparison
        ? referenceMapping.found && referenceMapping.mapping
          ? resolveMappedReference(comparisonSheet, referenceMapping.mapping, subModule)
          : { ok: false, note: `No comparison cell mapping found for ${subModule} / ${field} in the Executive reference sheet.` }
        : resolveSource(workbookData, source, iso, plantName);
      if (result.ok) {
        output.getCell(`G${rowNumber}`).value = result.text;
      }
      logs.push({
        sheet: sheetName,
        row: rowNumber,
        section: subModule,
        field,
        source,
        ...result,
      });
    }
  }

  addRunLog(outputWorkbook, logs);
  detachSharedFormulaFollowers(outputWorkbook);
  await outputWorkbook.xlsx.writeFile(outFile);
  return logs;
}

class ExecutiveDashboardValidator {
  constructor({ dgrFile, template, comparison, rules, dates, site, siteCfg, outFile }) {
    this.dgrFile = dgrFile;
    this.template = template;
    this.comparison = comparison;
    this.rules = rules;
    this.dates = dates;
    this.site = site;
    this.siteCfg = siteCfg;
    this.outFile = outFile;
  }

  async generate() {
    return generateExecutive(this);
  }
}

module.exports = {
  ExecutiveDashboardValidator,
  generateExecutive,
  resolveSource,
  findReferenceMapping,
  comparisonCellAddresses,
  mappedCellCoverage,
  selectComparisonSheet,
  resolveMappedReference,
  getMappedCellValue,
  findDateColumn,
  dateToIso,
  readDgrWorkbook,
};
