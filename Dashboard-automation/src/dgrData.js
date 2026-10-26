'use strict';
const XLSX = require('xlsx');
const { colToIndex, isoToSerial, serialToIso, isNum } = require('./utils');

/**
 * Loads only the sheets we need from the (large) DGR workbook and exposes
 * small helpers for looking up values by (sheet, column letter, date).
 */
class DgrData {
  constructor(file, sheetNames) {
    const wb = XLSX.readFile(file, { sheets: sheetNames, cellDates: false, cellFormula: false, cellNF: true, cellStyles: false });
    this.sheets = {};
    this.numberFormats = {};
    for (const name of sheetNames) {
      const ws = wb.Sheets[name];
      if (!ws) throw new Error(`Sheet "${name}" not found in ${file}. Found: ${wb.SheetNames.join(', ')}`);
      this.sheets[name] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
      this.numberFormats[name] = new Map();
      for (const address of Object.keys(ws)) {
        if (address.startsWith('!')) continue;
        const cell = ws[address];
        if (cell && cell.z) {
          const position = XLSX.utils.decode_cell(address);
          this.numberFormats[name].set(`${position.r}:${position.c}`, cell.z);
        }
      }
    }
  }

  rows(sheet) { return this.sheets[sheet]; }

  numberFormat(sheet, rowIndex, columnIndex) {
    return this.numberFormats[sheet]?.get(`${rowIndex}:${columnIndex}`);
  }

  /** all row-arrays of `sheet` whose `dateCol` falls on the ISO date (optionally filtered by plant) */
  rowsForDate(sheet, dateCol, iso, plantCol, plantName) {
    const target = isoToSerial(iso);
    const dc = colToIndex(dateCol);
    const pc = plantCol ? colToIndex(plantCol) : null;
    return this.sheets[sheet].filter((r, i) => {
      if (i === 0 || !isNum(r[dc]) || Math.floor(r[dc]) !== target) return false;
      if (plantName && pc !== null && String(r[pc] || '').trim().toLowerCase() !== plantName.toLowerCase()) return false;
      return true;
    });
  }

  availableDates(sheet, dateCol, plantCol, plantName) {
    const dc = colToIndex(dateCol), pc = plantCol ? colToIndex(plantCol) : null;
    const set = new Set();
    this.sheets[sheet].forEach((r, i) => {
      if (i === 0 || !isNum(r[dc])) return;
      if (plantName && pc !== null && String(r[pc] || '').trim().toLowerCase() !== plantName.toLowerCase()) return;
      set.add(serialToIso(r[dc]));
    });
    return [...set].sort();
  }
}

module.exports = { DgrData };
