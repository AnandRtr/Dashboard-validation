'use strict';

/** Parses "Expected:76.9139%", "5.60 (6.38)", "166.27 MWh" -> { num, unit, time } */
function parseValue(raw) {
  if (raw === null || raw === undefined) return null;
  let s = String(raw).trim();
  s = s.replace(/^Expected:\s*/i, '').replace(/^Actual:\s*/i, '');
  const timeMatch = s.match(/\(([\d.]+)\)\s*$/);
  const time = timeMatch ? timeMatch[1] : null;
  if (timeMatch) s = s.slice(0, timeMatch.index).trim();
  const numMatch = s.match(/-?[\d,]+(?:\.\d+)?/);
  if (!numMatch) return { num: null, unit: s, time };
  const num = Number(numMatch[0].replace(/,/g, ''));
  const unit = s.slice(numMatch.index + numMatch[0].length).trim();
  return { num, unit, time };
}

/** Compares Expected vs Actual text at `decimals` precision, +/- `tolerance`. */
function compareValues(expectedText, actualText, { decimals = 2, tolerance = 0 } = {}) {
  const exp = parseValue(expectedText);
  const act = parseValue(actualText);
  if (!exp || !act || exp.num === null || act.num === null) {
    return { pass: false, reason: 'Could not parse a number from Expected or Actual', exp, act };
  }
  const f = 10 ** decimals;
  const e = Math.round(exp.num * f) / f;
  const a = Math.round(act.num * f) / f;
  const pass = Math.abs(e - a) <= tolerance + 1e-9;
  return { pass, reason: pass ? 'Match' : `|${e} - ${a}| > ${tolerance}`, exp, act, expRounded: e, actRounded: a };
}

module.exports = { parseValue, compareValues };
