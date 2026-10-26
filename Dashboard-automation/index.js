#!/usr/bin/env node
'use strict';
const path = require('path');
const fs = require('fs');
const { parseArgs } = require('node:util');
const { DgrData } = require('./src/dgrData');
const { generate } = require('./src/generate');
const { ExecutiveDashboardValidator } = require('./src/executiveDashboard');

const HELP = `
Validation-dashboard automation
  node index.js --dgr <DGR.xlsx> --template <validation.xlsx> --comparison <KPI.xlsx> --site TG1.1 --date 2026-07-16 [--out out.xlsx] [--dashboard executive]

  --dgr        DGR workbook (has EXECUTIVE_DGR, Weather, Inverter_Wise, AC_Energy, Expected Power sheets)
  --template   Validation sheet template (Sub Module / Field Name / Formula / Comments layout)
  --comparison KPI comparison workbook for Executive mode (Comments tab with Sub Module, Field Name, and date columns)
  --site       Site label, e.g. TG1.1 (looked up in config/rules.json -> sites for the plant name)
  --date       YYYY-MM-DD (repeat or comma-separate for several days)
  --all-dates  Generate one sheet for every date present in the DGR workbook
  --dashboard  Validation type: standard or executive (default: standard)
  --out        Output file (default: ./report/<site>_Module-1_validation.xlsx)
`;

(async () => {
  const { values: a } = parseArgs({ options: {
    dgr: { type: 'string' }, template: { type: 'string' }, comparison: { type: 'string' }, site: { type: 'string' },
    date: { type: 'string', multiple: true }, 'all-dates': { type: 'boolean' }, dashboard: { type: 'string' },
    out: { type: 'string' }, help: { type: 'boolean' },
  } });
  if (a.help || !a.dgr || !a.template || (!a.date && !a['all-dates'])) { console.log(HELP); process.exit(a.help ? 0 : 1); }
  const mode = (a.dashboard || 'standard').toLowerCase();
  if (mode && !['standard', 'executive'].includes(mode)) { console.error('x --dashboard must be "standard" or "executive"'); process.exit(1); }
  if (mode === 'executive' && !a.comparison) {
    console.error('x Executive mode requires --comparison <KPI.xlsx>.');
    process.exit(1);
  }

  const rules = JSON.parse(fs.readFileSync(path.join(__dirname, 'config', 'rules.json'), 'utf8'));
  const siteCfg = rules.sites[a.site];
  if (a.site && !siteCfg) console.warn(`! Site "${a.site}" is not in config/rules.json -> sites; plant filter disabled.`);

  const standardSheetNames = [...new Set([
    ...Object.values(rules.sheetAliases).map((x) => x.sheet), rules.inverter.sheet,
    rules.specials.acEnergy.sheet, rules.specials.expectedPower.sheet])];
  const sheetNames = mode === 'executive'
    ? [rules.sheetAliases.exec_dgr.sheet]
    : standardSheetNames;
  console.log(`Reading ${path.basename(a.dgr)} (${sheetNames.join(', ')}) ...`);
  const data = new DgrData(a.dgr, sheetNames);

  const ex = rules.sheetAliases.exec_dgr;
  const avail = data.availableDates(ex.sheet, ex.dateCol, rules.plantNameColumn[ex.sheet], siteCfg && siteCfg.plantName);
  let dates = a['all-dates'] ? avail : (a.date || []).flatMap((d) => d.split(',')).map((d) => d.trim());
  const missing = dates.filter((d) => !avail.includes(d));
  if (missing.length) { console.error(`x No DGR data for: ${missing.join(', ')}\n  Available dates: ${avail.join(', ')}`); process.exit(2); }

  fs.mkdirSync('report', { recursive: true });
  const out = a.out || path.join('report', `${a.site || 'site'}_Module-1_validation.xlsx`);
  const validator = mode === 'executive'
    ? new ExecutiveDashboardValidator({ dgrFile: a.dgr, template: a.template, comparison: a.comparison, rules, dates, site: a.site, siteCfg, outFile: out })
    : { generate: () => generate({ template: a.template, data, rules, dates, site: a.site, siteCfg, outFile: out }) };
  const logs = await validator.generate();

  const bad = logs.filter((l) => !l.ok && !l.skip);
  console.log(`\nDone -> ${out}`);
  console.log(`Dates: ${dates.join(', ')} | rows filled: ${logs.filter((l) => l.ok).length} | skipped: ${logs.filter((l) => l.skip).length} | errors: ${bad.length}`);
  bad.slice(0, 20).forEach((l) => console.log(`  [${l.sheet}] row ${l.row} ${l.section} / ${l.field}: ${l.note}`));
})().catch((e) => { console.error(e); process.exit(1); });
