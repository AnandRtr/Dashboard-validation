'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { DgrData } = require('./src/dgrData');
const { generate } = require('./src/generate');
const { ExecutiveDashboardValidator } = require('./src/executiveDashboard');
const { AppReader } = require('./src/appReader');
const { colToIndex } = require('./src/utils');
const { compareValues } = require('./src/compare');

const app = express();
const port = Number(process.env.PORT) || 3000;
const root = __dirname;
const rules = JSON.parse(fs.readFileSync(path.join(root, 'config', 'rules.json'), 'utf8'));
const apiFieldsPath = path.join(root, 'config', 'apiFields.json');
const apiFields = fs.existsSync(apiFieldsPath) ? JSON.parse(fs.readFileSync(apiFieldsPath, 'utf8')) : null;
const appLabelsPath = path.join(root, 'config', 'appLabels.json');
const appLabels = JSON.parse(fs.readFileSync(appLabelsPath, 'utf8'));
const reportDir = path.join(root, 'report');
const authPath = path.join(root, 'auth.json');
const validationTemplates = {
  standard: path.join(root, 'inputs', 'Dashboard_validation_sheet.xlsx'),
  executive: path.join(root, 'inputs', 'Executive-Dashboard-reference.xlsx'),
};
const DEFAULT_APP_URL = process.env.APP_URL || 'https://dxt-azure-dev.netlify.app';
const GREEN = 'FF92D050', RED = 'FFFF0000';
const uploadDir = fs.mkdtempSync(path.join(os.tmpdir(), 'validation-dashboard-'));
const upload = multer({
  storage: multer.diskStorage({
    destination: uploadDir,
    filename: (_req, file, callback) => callback(null, `${crypto.randomUUID()}${path.extname(file.originalname).toLowerCase()}`),
  }),
  limits: { fileSize: 100 * 1024 * 1024, files: 2, fields: 5 },
  fileFilter: (_req, file, callback) => {
    const extension = path.extname(file.originalname).toLowerCase();
    const allowed = file.fieldname === 'dgr' ? ['.xlsx', '.xls'] : ['.xlsx'];
    callback(allowed.includes(extension) ? null : new Error(`Unsupported file type for ${file.fieldname}.`), allowed.includes(extension));
  },
});

const sheetNames = [...new Set([
  ...Object.values(rules.sheetAliases).map((alias) => alias.sheet),
  rules.inverter.sheet,
  rules.specials.acEnergy.sheet,
  rules.specials.expectedPower.sheet,
])];

function loadDgr(filePath, selectedSite, dashboard = 'standard') {
  const requestedSheets = dashboard === 'executive'
    ? [rules.sheetAliases.exec_dgr.sheet]
    : sheetNames;
  const data = new DgrData(filePath, requestedSheets);
  const execSheet = rules.sheetAliases.exec_dgr;
  const plantColumn = colToIndex(rules.plantNameColumn[execSheet.sheet]);
  const plantNames = [...new Set(data.rows(execSheet.sheet)
    .slice(1)
    .map((row) => String(row[plantColumn] || '').trim())
    .filter(Boolean))];
  const sites = plantNames.map((plantName) => {
    const alias = Object.entries(rules.sites).find(([, config]) => config.plantName.toLowerCase() === plantName.toLowerCase());
    return {
      value: alias ? alias[0] : plantName,
      label: plantName,
      plantName,
      dates: data.availableDates(execSheet.sheet, execSheet.dateCol, rules.plantNameColumn[execSheet.sheet], plantName),
    };
  }).filter((site) => site.dates.length > 0);
  const site = sites.find((option) => option.value === selectedSite);
  return { data, sites, site, dates: site ? site.dates : [], siteCfg: site ? { plantName: site.plantName } : undefined };
}

function cleanup(files) {
  for (const file of files) {
    if (file) fs.rm(file.path, { force: true }, () => {});
  }
}

function requireSite(site, availableSites) {
  if (!site || !availableSites.some((option) => option.value === site)) {
    const error = new Error('Choose a plant found in the uploaded DGR workbook.');
    error.status = 400;
    throw error;
  }
}

function normalizeDate(value) {
  let isoDate = value;
  if (/^\d{2}-\d{2}-\d{4}$/.test(value)) {
    const [day, month, year] = value.split('-');
    isoDate = `${year}-${month}-${day}`;
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return null;
  const parsed = new Date(`${isoDate}T00:00:00.000Z`);
  return Number.isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== isoDate ? null : isoDate;
}

app.use(express.static(path.join(root, 'public')));
app.use(express.json());

app.get('/api/config', (_req, res) => {
  res.json({ sites: Object.keys(rules.sites), stations: Object.keys(apiFields?.stations || {}) });
});

app.get('/api/reports', (_req, res) => {
  fs.mkdirSync(reportDir, { recursive: true });
  const reports = fs.readdirSync(reportDir)
    .filter((name) => name.toLowerCase().endsWith('.xlsx'))
    .map((name) => {
      const stats = fs.statSync(path.join(reportDir, name));
      return { name, size: stats.size, createdAt: stats.mtime.toISOString() };
    })
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  res.json(reports);
});

app.post('/api/inspect', upload.single('dgr'), (req, res, next) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Upload a DGR workbook first.' });
    const dashboard = req.body.dashboard || 'standard';
    if (!['standard', 'executive'].includes(dashboard)) {
      return res.status(400).json({ error: 'Choose a valid dashboard type.' });
    }
    const { sites, dates } = loadDgr(req.file.path, req.body.site, dashboard);
    return res.json({ sites: sites.map(({ value, label, dates: siteDates }) => ({ value, label, dates: siteDates })), dates });
  } catch (error) {
    return next(error);
  } finally {
    cleanup([req.file]);
  }
});

app.post('/api/generate', upload.fields([
  { name: 'dgr', maxCount: 1 },
  { name: 'comparison', maxCount: 1 },
]), async (req, res, next) => {
  const files = Object.values(req.files || {}).flat();
  try {
    const dgr = req.files && req.files.dgr && req.files.dgr[0];
    const comparison = req.files && req.files.comparison && req.files.comparison[0];
    if (!dgr) return res.status(400).json({ error: 'Upload the DGR workbook.' });

    const dashboard = req.body.dashboard || 'standard';
    if (!['standard', 'executive'].includes(dashboard)) {
      return res.status(400).json({ error: 'Choose a valid dashboard type.' });
    }
    const template = validationTemplates[dashboard];
    if (!fs.existsSync(template)) {
      return res.status(500).json({ error: `The ${dashboard} validation reference is missing: ${path.basename(template)}.` });
    }
    if (dashboard === 'executive' && !comparison) {
      return res.status(400).json({ error: 'Upload the KPI comparison workbook for Executive Dashboard validation.' });
    }
    const { data, sites, site, dates: availableDates, siteCfg } = loadDgr(dgr.path, req.body.site, dashboard);
    requireSite(req.body.site, sites);
    const allDates = req.body.allDates === 'true';
    let dates;
    try {
      dates = allDates ? availableDates : JSON.parse(req.body.dates || '[]');
    } catch {
      return res.status(400).json({ error: 'The selected dates could not be read.' });
    }
    if (!Array.isArray(dates) || dates.length === 0) {
      return res.status(400).json({ error: 'Select at least one date, or choose all available dates.' });
    }
    const invalidDates = dates.filter((date) => !availableDates.includes(date));
    if (invalidDates.length) return res.status(400).json({ error: `No DGR data for: ${invalidDates.join(', ')}.` });

    fs.mkdirSync(reportDir, { recursive: true });
    const siteSlug = req.body.site.replace(/[^a-z0-9.-]+/gi, '_');
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const dashboardSuffix = dashboard === 'executive' ? 'executive_validation' : 'validation';
    const filename = `${siteSlug}_${dashboardSuffix}_${stamp}.xlsx`;
    const outFile = path.join(reportDir, filename);
    const logs = dashboard === 'executive'
      ? await new ExecutiveDashboardValidator({
        dgrFile: dgr.path,
        template,
        comparison: comparison.path,
        rules,
        dates,
        site: req.body.site,
        siteCfg,
        outFile,
      }).generate()
      : await generate({ template, data, rules, dates, site: req.body.site, siteCfg, outFile });
    const summary = {
      filename,
      dates,
      filled: logs.filter((row) => row.ok).length,
      skipped: logs.filter((row) => row.skip).length,
      errors: logs.filter((row) => !row.ok && !row.skip).length,
      issues: logs.filter((row) => !row.ok).slice(0, 80).map((row) => ({
        sheet: row.sheet,
        row: row.row,
        section: row.section,
        field: row.field,
        skipped: Boolean(row.skip),
        note: row.note,
      })),
    };
    return res.json(summary);
  } catch (error) {
    return next(error);
  } finally {
    cleanup(files);
  }
});

app.get('/api/reports/:name/sheets', async (req, res, next) => {
  const filename = path.basename(req.params.name);
  if (filename !== req.params.name || !filename.toLowerCase().endsWith('.xlsx')) {
    return res.status(404).json({ error: 'Report not found.' });
  }
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(reportDir, filename));
    const sheets = wb.worksheets.filter((ws) => ws.name !== 'Run_Log').map((ws) => ws.name);
    return res.json({ sheets });
  } catch (error) {
    return next(error);
  }
});

app.delete('/api/reports/:name', async (req, res, next) => {
  const filename = path.basename(req.params.name);
  if (filename !== req.params.name || !filename.toLowerCase().endsWith('.xlsx')) {
    return res.status(404).json({ error: 'Report not found.' });
  }
  const reportPath = path.join(reportDir, filename);
  try {
    await fs.promises.unlink(reportPath);
    return res.status(204).end();
  } catch (error) {
    if (error.code === 'ENOENT') return res.status(404).json({ error: 'Report not found.' });
    return next(error);
  }
});

app.post('/api/validate', async (req, res, next) => {
  let reader;
  try {
    const { report, sheet, station, date } = req.body || {};
    if (!report || !sheet || !station || !date) {
      return res.status(400).json({ error: 'report, sheet, station and date are all required.' });
    }
    const isoDate = normalizeDate(date);
    if (!isoDate) return res.status(400).json({ error: 'Date must be YYYY-MM-DD or DD-MM-YYYY.' });
    const filename = path.basename(report);
    const filePath = path.join(reportDir, filename);
    if (filename !== report || !fs.existsSync(filePath)) return res.status(404).json({ error: 'Report not found.' });
    if (!fs.existsSync(authPath)) {
      return res.status(400).json({ error: 'auth.json not found next to server.js. Refresh it with: npx playwright codegen --save-storage=auth.json <app login url>' });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet(sheet);
    if (!ws) return res.status(404).json({ error: `Sheet "${sheet}" not found in ${filename}.` });
    const sheetContext = /^(.*)_(\d{4}-\d{2}-\d{2})$/.exec(sheet);
    if (sheetContext && isoDate !== sheetContext[2]) {
      return res.status(400).json({ error: `Date must match the selected report sheet (${sheetContext[2]}).` });
    }

    if (!fs.existsSync(authPath)) {
      return res.status(400).json({ error: 'auth.json not found next to server.js. Sign in and refresh it with: npx playwright codegen --save-storage=auth.json https://dxt-azure-dev.netlify.app/login' });
    }

    reader = new AppReader({
      baseUrl: process.env.APP_URL || DEFAULT_APP_URL,
      storageState: authPath,
      headless: req.body.headless !== false,
    });
    await reader.open();
    await reader.selectStation(station);
    await reader.selectDate(isoDate);
    const actuals = await reader.readAll(appLabels);

    ws.getCell('G2').value = 'Actual';
    ws.getColumn(7).width = Math.max(ws.getColumn(7).width || 0, 30);
    ws.getCell('G2').style = JSON.parse(JSON.stringify(ws.getCell('F2').style));

    let subModule = '';
    const results = [];

    for (let r = 3; r <= ws.rowCount; r++) {
      const aCell = ws.getCell(`A${r}`);
      if (!(aCell.isMerged && aCell.master.address !== aCell.address) && aCell.value) {
        subModule = typeof aCell.value === 'object' ? aCell.value.richText.map((t) => t.text).join('') : String(aCell.value);
      }
      const field = String(ws.getCell(`B${r}`).value ?? '').trim();
      const expectedCell = ws.getCell(`F${r}`);
      let expected = expectedCell.value;
      if (typeof expected === 'string' && expected.includes(' Actual:')) {
        expected = expected.split(' Actual:', 1)[0];
        expectedCell.value = expected;
      }
      if (!field || !expected) continue;

      const cfg = appLabels[`${subModule}|${field}`];
      if (!cfg) continue;
      const rawActual = actuals[`${subModule}|${field}`];

      if (rawActual === undefined || rawActual === null) {
        ws.getCell(`G${r}`).style = JSON.parse(JSON.stringify(expectedCell.style));
        ws.getCell(`G${r}`).value = 'Not readable';
        results.push({ row: r, section: subModule, field, status: 'NO_DATA', note: `Could not read "${cfg.label}" from the dashboard` });
        continue;
      }

      const actualText = String(rawActual);
      const cmp = compareValues(String(expected), actualText, { decimals: cfg.decimals ?? 2 });
      ws.getCell(`G${r}`).style = JSON.parse(JSON.stringify(expectedCell.style));
      ws.getCell(`G${r}`).value = actualText;
      ws.getCell(`E${r}`).value = cmp.pass ? 'Pass' : 'Fail';
      ws.getCell(`E${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cmp.pass ? GREEN : RED } };
      results.push({ row: r, section: subModule, field, status: cmp.pass ? 'PASS' : 'FAIL', note: cmp.reason, expected, actual: actualText });
    }

    await wb.xlsx.writeFile(filePath);
    return res.json({
      filename,
      sheet,
      pass: results.filter((r) => r.status === 'PASS').length,
      fail: results.filter((r) => r.status === 'FAIL').length,
      noData: results.filter((r) => r.status === 'NO_DATA').length,
      issues: results.filter((r) => r.status !== 'PASS'),
    });
  } catch (error) {
    return next(error);
  } finally {
    if (reader) await reader.close().catch(() => {});
  }
});

app.get('/api/reports/:name', (req, res, next) => {
  const filename = path.basename(req.params.name);
  if (filename !== req.params.name || !filename.toLowerCase().endsWith('.xlsx')) {
    return res.status(404).json({ error: 'Report not found.' });
  }
  return res.download(path.join(reportDir, filename), filename, (error) => {
    if (error && !res.headersSent) next(error);
  });
});

app.use((error, _req, res, _next) => {
  const status = error.status || (error instanceof multer.MulterError ? 400 : 422);
  res.status(status).json({ error: error.message || 'The request could not be completed.' });
});

app.listen(port, () => {
  console.log(`Validation dashboard listening at http://localhost:${port}`);
});

process.on('exit', () => fs.rmSync(uploadDir, { recursive: true, force: true }));