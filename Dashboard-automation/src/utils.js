'use strict';

/** 'AM' -> 38 (0-based) */
function colToIndex(col) {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** 'YYYY-MM-DD' -> Excel serial day number (integer) */
function isoToSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) throw new Error(`Bad date "${iso}". Use YYYY-MM-DD`);
  return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569;
}

function serialToIso(serial) {
  const d = new Date(Math.round((Math.floor(serial) - 25569) * 86400000));
  return d.toISOString().slice(0, 10);
}

/** fraction-of-day serial -> {h, m} */
function serialToTime(serial) {
  const mins = Math.round((serial - Math.floor(serial)) * 1440);
  return { h: Math.floor(mins / 60) % 24, m: mins % 60 };
}

/** 18:38 -> "6.38" (the 12h "h.mm" style used in the existing validation sheet) */
function fmtTime12(serial) {
  const { h, m } = serialToTime(serial);
  return `${h % 12 || 12}.${String(m).padStart(2, '0')}`;
}

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);

/** Preserve the source number exactly; validation reports must not round or truncate values. */
function formatNumber(value) {
  return String(value);
}

function percentValue(value) {
  const match = /^([+-]?)(\d+)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(String(value));
  if (!match) return String(value * 100);
  const [, sign, whole, fraction = '', exponent = '0'] = match;
  const digits = `${whole}${fraction}`;
  const decimalPosition = whole.length + Number(exponent) + 2;
  let scaled;
  if (decimalPosition <= 0) {
    scaled = `0.${'0'.repeat(-decimalPosition)}${digits}`;
  } else if (decimalPosition >= digits.length) {
    scaled = `${digits}${'0'.repeat(decimalPosition - digits.length)}`;
  } else {
    scaled = `${digits.slice(0, decimalPosition)}.${digits.slice(decimalPosition)}`;
  }
  if (scaled.includes('.')) scaled = scaled.replace(/0+$/, '').replace(/\.$/, '');
  const [integer, decimals] = scaled.split('.');
  const normalizedInteger = integer.replace(/^0+(?=\d)/, '');
  return `${sign}${normalizedInteger || '0'}${decimals ? `.${decimals}` : ''}`;
}

function formatSourceValue(value, numberFormat) {
  if (!isNum(value)) return String(value);
  return numberFormat && /%/.test(numberFormat) ? `${percentValue(value)}%` : formatNumber(value);
}

function copyCellValue(cell) {
  const value = cell.value;
  if (!value || typeof value !== 'object') return value;
  if (Object.prototype.hasOwnProperty.call(value, 'sharedFormula')) {
    if (value.result === undefined) {
      throw new Error(`Cannot copy shared formula cell ${cell.address} without a cached result.`);
    }
    return value.result;
  }
  if (Object.prototype.hasOwnProperty.call(value, 'formula')) {
    return { formula: value.formula, result: value.result };
  }
  return value;
}

function detachSharedFormulaFollowers(workbook) {
  const detached = [];
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const value = cell.value;
        if (!value || typeof value !== 'object' || !Object.prototype.hasOwnProperty.call(value, 'sharedFormula')) return;

        if (value.result !== undefined) {
          cell.value = value.result;
        } else if (cell.formula) {
          cell.value = { formula: cell.formula };
        } else {
          throw new Error(`Cannot write ${worksheet.name}!${cell.address}: shared formula ${JSON.stringify(value)} has no master or cached value.`);
        }
        detached.push(`${worksheet.name}!${cell.address}`);
      });
    });
  }
  return detached;
}

function runLogValue(value) {
  if (value === null || value === undefined || typeof value !== 'object' || value instanceof Date) return value;
  if (value.result !== undefined) return value.result;
  if (value.sharedFormula) return `Shared formula reference: ${value.sharedFormula}`;
  if (value.formula) return `Formula: ${value.formula}`;
  if (value.richText) return value.richText.map((part) => part.text).join('');
  return JSON.stringify(value);
}

module.exports = {
  colToIndex,
  isoToSerial,
  serialToIso,
  serialToTime,
  fmtTime12,
  isNum,
  formatNumber,
  formatSourceValue,
  copyCellValue,
  detachSharedFormulaFollowers,
  runLogValue,
};
