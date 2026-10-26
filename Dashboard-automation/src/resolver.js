'use strict';
const { colToIndex, isNum, formatSourceValue, fmtTime12 } = require('./utils');

/**
 * Resolves the "Expected" text for one validation row.
 * Returns { ok, text, raw, ref, note }.
 */
class Resolver {
  constructor(data, rules, { iso, plantName }) {
    this.data = data; this.rules = rules; this.iso = iso; this.plantName = plantName;
    this.cache = new Map();
    const keys = Object.keys(rules.sheetAliases).sort((a, b) => b.length - a.length).map((k) => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    this.commentRe = new RegExp(`^\\s*(${keys.join('|')})_([A-Za-z]{1,3})(?![A-Za-z0-9])`, 'i');
  }

  fmtOpts(rule) { return { ...this.rules.defaults, ...rule }; }

  sourceValue(sheet, row, column, value) {
    const rowIndex = this.data.rows(sheet).indexOf(row);
    const numberFormat = rowIndex < 0 ? null : this.data.numberFormat(sheet, rowIndex, column);
    return formatSourceValue(value, numberFormat);
  }

  wrap(rule, body) {
    const o = this.fmtOpts(rule);
    return (o.prefix ? 'Expected:' : '') + body;
  }

  dayRows(sheet, dateCol, plantCol) {
    const key = `${sheet}|${dateCol}`;
    if (!this.cache.has(key)) {
      const pn = plantCol ? this.plantName : null;
      this.cache.set(key, this.data.rowsForDate(sheet, dateCol, this.iso, plantCol, pn));
    }
    return this.cache.get(key);
  }

  /** value from a "Sheet_COL" reference */
  fromCell(alias, col, rule) {
    const { sheet, dateCol } = alias;
    const plantCol = this.rules.plantNameColumn[sheet];
    const rows = this.dayRows(sheet, dateCol, plantCol);
    if (!rows.length) return { ok: false, note: `No row for ${this.iso} in ${sheet}` };
    const raw = rows[0][colToIndex(col)];
    const ref = `${sheet}!${col.toUpperCase()}`;
    if (raw === null || raw === undefined || raw === '') return { ok: false, ref, note: `${ref} is empty` };
    if (!isNum(raw)) return { ok: true, ref, raw, text: this.wrap(rule, String(raw)) };
    return { ok: true, ref, raw, text: this.wrap(rule, this.sourceValue(sheet, rows[0], colToIndex(col), raw)) };
  }

  /** returns {time, ac, gti, exp} entries of the day from the time-series sheets */
  special(name, rule) {
    const sp = this.rules.specials;
    const thr = sp.gtiThreshold;

    if (name === 'acLastActive') {
      const c = sp.acEnergy;
      const rows = this.dayRows(c.sheet, c.dateCol);
      const vi = colToIndex(c.valueCol), ti = colToIndex(c.timeCol);
      const act = rows.filter((r) => isNum(r[vi]) && r[vi] > 0);
      if (!act.length) return { ok: false, note: `No active (>0) ${c.sheet}!${c.valueCol} value on ${this.iso}` };
      const last = act[act.length - 1];
      return { ok: true, ref: `${c.sheet}!${c.valueCol} (last >0)`, raw: last[vi], text: this.wrap(rule, `${this.sourceValue(c.sheet, last, vi, last[vi])} (${fmtTime12(last[ti])})`) };
    }

    const c = sp.expectedPower;
    const rows = this.dayRows(c.sheet, c.dateCol);
    const gi = colToIndex(c.gtiCol), vi = colToIndex(c.valueCol), ti = colToIndex(c.timeCol);
    if (name === 'peakGti') {
      const g = rows.filter((r) => isNum(r[gi]));
      if (!g.length) return { ok: false, note: `No GTI in ${c.sheet}` };
      const max = Math.max(...g.map((r) => r[gi]));
      const maxRow = g.reduce((best, row) => row[gi] > best[gi] ? row : best);
      return { ok: true, ref: `${c.sheet}!${c.gtiCol} (max)`, raw: max, text: this.wrap(rule, this.sourceValue(c.sheet, maxRow, gi, max)) };
    }
    const last = rows.filter((r) => isNum(r[gi]) && r[gi] > thr).pop();
    if (!last) return { ok: false, note: `No GTI > ${thr} on ${this.iso}` };
    if (name === 'gtiLastGt5') {
      return { ok: true, ref: `${c.sheet}!${c.gtiCol} (last >${thr})`, raw: last[gi], text: this.wrap(rule, `${this.sourceValue(c.sheet, last, gi, last[gi])} (${fmtTime12(last[ti])})`) };
    }
    if (name === 'expectedPowerLastGt5') {
      if (!isNum(last[vi])) return { ok: false, note: `${c.sheet}!${c.valueCol} empty at last GTI>${thr}` };
      return { ok: true, ref: `${c.sheet}!${c.valueCol} (at last GTI>${thr})`, raw: last[vi], text: this.wrap(rule, `${this.sourceValue(c.sheet, last, vi, last[vi])} (${fmtTime12(last[ti])})`) };
    }
    return { ok: false, note: `Unknown special "${name}"` };
  }

  /** Top-of-sheet rows (Performance KPI, Weather, ...) */
  resolveField(subModule, field, comment) {
    const rule = this.rules.fields[`${subModule}|${field}`];
    if (!rule) return { ok: false, skip: true, note: 'No rule configured for this row (add it to config/rules.json)' };
    if (rule.special) return this.special(rule.special, rule);
    if (rule.source) {
      const alias = Object.values(this.rules.sheetAliases).find((a) => a.sheet === rule.source.sheet);
      return this.fromCell(alias, rule.source.col, rule);
    }
    const m = comment ? this.commentRe.exec(String(comment)) : null;
    if (!m) return { ok: false, skip: !comment, note: comment ? `Cannot parse source from comment "${comment}"` : 'Comment cell is empty' };
    const alias = this.rules.sheetAliases[m[1].toLowerCase()];
    return this.fromCell(alias, m[2].toUpperCase(), rule);
  }

  /** Inverter rows: value of KPI for INVn from Inverter_Wise */
  resolveInverter(inv, field) {
    const cfg = this.rules.inverter;
    const rule = cfg.fields[field];
    if (!rule) return { ok: false, skip: true, note: `No inverter rule for "${field}"` };
    if (rule.skip) return { ok: false, skip: true, note: rule.skip };
    const rows = this.dayRows(cfg.sheet, cfg.dateCol);
    const header = this.data.rows(cfg.sheet)[0];
    const normalizeInv = (value) => String(value).replace(/[^A-Z0-9]/gi, '').toUpperCase();
    const ci = header.findIndex((h, i) => i >= colToIndex(cfg.firstInverterCol) && normalizeInv(h) === normalizeInv(inv));
    if (ci < 0) return { ok: false, note: `${inv} column not in ${cfg.sheet}` };
    const kpiIdx = colToIndex(cfg.kpiCol);
    const row = rows.find((r) => String(r[kpiIdx]).trim() === rule.kpi);
    if (!row) return { ok: false, note: `KPI "${rule.kpi}" not found for ${this.iso}` };
    const raw = row[ci];
    if (!isNum(raw)) return { ok: false, note: `${inv} ${rule.kpi} empty` };
    return { ok: true, raw, ref: `${cfg.sheet}!${rule.kpi}/${inv}`, text: this.wrap(rule, this.sourceValue(cfg.sheet, row, ci, raw)) };
  }

  /** is inverter inactive for the day? (INV_GEN = 0) */
  isInverterInactive(inv) {
    const cfg = this.rules.inverter;
    const rows = this.dayRows(cfg.sheet, cfg.dateCol);
    const header = this.data.rows(cfg.sheet)[0];
    const normalizeInv = (value) => String(value).replace(/[^A-Z0-9]/gi, '').toUpperCase();
    const ci = header.findIndex((h, i) => i >= colToIndex(cfg.firstInverterCol) && normalizeInv(h) === normalizeInv(inv));
    const row = rows.find((r) => String(r[colToIndex(cfg.kpiCol)]).trim() === cfg.inactiveWhenKpi);
    return !!row && ci >= 0 && row[ci] === 0;
  }
}

module.exports = { Resolver };
