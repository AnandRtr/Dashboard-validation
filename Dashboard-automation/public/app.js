'use strict';

const siteSelect = document.querySelector('#site-select');
const dashboardSelect = document.querySelector('#dashboard-select');
const dgrInput = document.querySelector('#dgr-file');
const comparisonInput = document.querySelector('#comparison-file');
const dateList = document.querySelector('#date-list');
const allDatesInput = document.querySelector('#all-dates');
const runButton = document.querySelector('#run-button');
const notice = document.querySelector('#notice');
const resultPanel = document.querySelector('#result-panel');
const reportBody = document.querySelector('#reports-body');
const reportCount = document.querySelector('#report-count');
const dateHelp = document.querySelector('#date-help');
const emptyDates = document.querySelector('#empty-dates');
let availableDates = [];
let availableSites = [];
let inspecting = false;
let generating = false;

function setNotice(message, kind = 'info') {
  notice.textContent = message;
  notice.classList.toggle('error', kind === 'error');
  notice.hidden = !message;
}

async function requestJson(url, options) {
  const response = await fetch(url, options);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || `Request failed (${response.status}).`);
  return payload;
}

function updateRunButton() {
  const hasDates = allDatesInput.checked ? availableDates.length > 0 : dateList.querySelectorAll('input:checked').length > 0;
  const comparisonRequired = dashboardSelect.value === 'executive';
  runButton.disabled = generating || inspecting || !dgrInput.files.length
    || (comparisonRequired && !comparisonInput.files.length) || !siteSelect.value || !hasDates;
}

function showFile(input, labelId, dropzoneId) {
  const file = input.files[0];
  const card = document.querySelector(`#${dropzoneId}`);
  const label = document.querySelector(`#${labelId}`);
  label.firstChild.textContent = file ? file.name : 'Choose a file ';
  card.classList.toggle('has-file', Boolean(file));
  card.title = file ? file.name : '';
}

function renderDates(dates) {
  availableDates = dates;
  dateList.replaceChildren(...dates.map((date) => {
    const label = document.createElement('label');
    label.className = 'date-option';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = date;
    checkbox.disabled = allDatesInput.checked;
    const text = document.createElement('span');
    text.textContent = date;
    label.append(checkbox, text);
    return label;
  }));
  dateList.classList.toggle('is-disabled', dates.length === 0 || allDatesInput.checked);
  allDatesInput.disabled = dates.length === 0;
  emptyDates.hidden = dates.length > 0;
  dateHelp.textContent = dates.length
    ? `${dates.length} date${dates.length === 1 ? '' : 's'} found. Use Ctrl or Command to select multiple.`
    : siteSelect.value ? 'No dated rows found for this plant in the DGR.' : 'Choose a plant from the uploaded DGR to load its dates.';
  updateRunButton();
}

async function inspectDgr() {
  const previousSite = siteSelect.value;
  availableSites = [];
  siteSelect.replaceChildren(new Option('Upload a DGR workbook first', ''));
  siteSelect.disabled = true;
  renderDates([]);
  if (!dgrInput.files.length) {
    dateHelp.textContent = 'Upload a DGR workbook to detect available dates.';
    return;
  }
  inspecting = true;
  updateRunButton();
  dateHelp.textContent = 'Reading available dates from DGR...';
  setNotice('Checking workbook sheets and matching dates...');
  const form = new FormData();
  form.append('dgr', dgrInput.files[0]);
  form.append('dashboard', dashboardSelect.value);
  try {
    const result = await requestJson('/api/inspect', { method: 'POST', body: form });
    availableSites = result.sites;
    siteSelect.replaceChildren(new Option(availableSites.length ? 'Select a plant site' : 'No plants with dated rows found', ''));
    for (const site of availableSites) siteSelect.add(new Option(site.label, site.value));
    siteSelect.disabled = availableSites.length === 0;
    const selectedSite = availableSites.find((site) => site.value === previousSite) || (availableSites.length === 1 ? availableSites[0] : null);
    siteSelect.value = selectedSite ? selectedSite.value : '';
    const selectedDates = selectedSite ? selectedSite.dates : [];
    renderDates(selectedDates);
    if (selectedSite && selectedDates.length) {
      setNotice(`${selectedSite.label}: ${selectedDates.length} available date${selectedDates.length === 1 ? '' : 's'}.`, 'info');
    } else if (availableSites.length) {
      setNotice(`${availableSites.length} plant${availableSites.length === 1 ? '' : 's'} found in the DGR. Choose one to load its dates.`);
    } else {
      setNotice('No plant with dated rows was found. Check the plant and date columns in the DGR workbook.', 'error');
    }
  } catch (error) {
    availableSites = [];
    siteSelect.replaceChildren(new Option('Could not read DGR plants', ''));
    renderDates([]);
    setNotice(error.message, 'error');
  } finally {
    inspecting = false;
    updateRunButton();
  }
}

function renderSummary(summary) {
  resultPanel.hidden = false;
  document.querySelector('#result-title').textContent = `${summary.filename}`;
  const download = document.querySelector('#result-download');
  download.href = `/api/reports/${encodeURIComponent(summary.filename)}`;
  download.setAttribute('download', summary.filename);
  const stats = [
    { value: summary.filled, label: 'Values filled' },
    { value: summary.skipped, label: 'Rows skipped' },
    { value: summary.errors, label: 'Needs attention', error: summary.errors > 0 },
  ];
  document.querySelector('#result-stats').replaceChildren(...stats.map((stat) => {
    const node = document.createElement('div');
    node.className = `stat${stat.error ? ' error' : ''}`;
    const count = document.createElement('strong');
    count.textContent = String(stat.value);
    const label = document.createElement('span');
    label.textContent = stat.label;
    node.append(count, label);
    return node;
  }));
  const issueList = document.querySelector('#issue-list');
  issueList.replaceChildren(...summary.issues.map((issue) => {
    const row = document.createElement('div');
    row.className = `issue-row${issue.skipped ? ' skipped' : ''}`;
    const symbol = document.createElement('span');
    symbol.className = 'issue-symbol';
    symbol.textContent = issue.skipped ? '–' : '!';
    const text = document.createElement('span');
    text.textContent = `${issue.sheet} · row ${issue.row} · ${issue.section || ''} / ${issue.field || ''}: ${issue.note || 'No value resolved.'}`;
    row.append(symbol, text);
    return row;
  }));
  resultPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function formatSize(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

async function loadReports() {
  const reports = await requestJson('/api/reports');
  reportCount.textContent = String(reports.length);
  document.querySelector('#reports-empty').hidden = reports.length > 0;
  reportBody.replaceChildren(...reports.map((report) => {
    const row = document.createElement('tr');
    const name = document.createElement('td');
    name.textContent = report.name;
    name.title = report.name;
    const created = document.createElement('td');
    created.textContent = new Date(report.createdAt).toLocaleString();
    const size = document.createElement('td');
    size.textContent = formatSize(report.size);
    const actionCell = document.createElement('td');
    const actions = document.createElement('div');
    actions.className = 'report-actions';
    const download = document.createElement('a');
    download.className = 'report-download';
    download.href = `/api/reports/${encodeURIComponent(report.name)}`;
    download.download = report.name;
    download.textContent = 'Download';
    const remove = document.createElement('button');
    remove.className = 'report-delete';
    remove.type = 'button';
    remove.dataset.report = report.name;
    remove.textContent = 'Delete';
    remove.setAttribute('aria-label', `Delete ${report.name}`);
    actions.append(download, remove);
    actionCell.append(actions);
    row.append(name, created, size, actionCell);
    return row;
  }));
}

async function loadConfig() {
  siteSelect.replaceChildren(new Option('Upload a DGR workbook first', ''));
}

document.querySelector('#today-label').textContent = new Intl.DateTimeFormat(undefined, { dateStyle: 'medium' }).format(new Date());
dashboardSelect.addEventListener('change', () => {
  const executive = dashboardSelect.value === 'executive';
  document.querySelector('#run-form').classList.toggle('executive-mode', executive);
  document.querySelector('#comparison-dropzone').hidden = !executive;
  comparisonInput.required = executive;
  document.querySelector('#template-file-name').textContent = executive
    ? 'Executive-Dashboard-reference.xlsx'
    : 'Dashboard_validation_sheet.xlsx';
  document.querySelector('#template-copy').textContent = executive
    ? 'Automatically selected for Executive Dashboard validation'
    : 'Automatically selected for Dashboard validation';
  setNotice(executive
    ? 'The Executive reference sheet is selected automatically. Upload the KPI comparison workbook for expected KPI/chart values.'
    : '');
  updateRunButton();
  inspectDgr();
});
dgrInput.addEventListener('change', () => {
  showFile(dgrInput, 'dgr-file-name', 'dgr-dropzone');
  inspectDgr();
});
comparisonInput.addEventListener('change', () => {
  showFile(comparisonInput, 'comparison-file-name', 'comparison-dropzone');
  updateRunButton();
});
siteSelect.addEventListener('change', () => {
  const site = availableSites.find((option) => option.value === siteSelect.value);
  renderDates(site ? site.dates : []);
  if (site && site.dates.length) setNotice(`${site.label}: ${site.dates.length} available date${site.dates.length === 1 ? '' : 's'}.`, 'info');
});
dateList.addEventListener('change', updateRunButton);
allDatesInput.addEventListener('change', () => {
  dateList.classList.toggle('is-disabled', allDatesInput.checked || availableDates.length === 0);
  dateList.querySelectorAll('input').forEach((checkbox) => { checkbox.disabled = allDatesInput.checked; });
  updateRunButton();
});
document.querySelector('#refresh-reports').addEventListener('click', () => loadReports().catch((error) => setNotice(error.message, 'error')));
reportBody.addEventListener('click', async (event) => {
  const button = event.target.closest('.report-delete');
  if (!button || !window.confirm(`Delete generated report "${button.dataset.report}"? This cannot be undone.`)) return;

  button.disabled = true;
  try {
    const reportName = button.dataset.report;
    await requestJson(`/api/reports/${encodeURIComponent(reportName)}`, { method: 'DELETE' });
    const latestDownload = document.querySelector('#result-download');
    if (latestDownload.href.endsWith(`/api/reports/${encodeURIComponent(reportName)}`)) {
      resultPanel.hidden = true;
    }
    await loadReports();
    setNotice(`Deleted report "${reportName}".`);
  } catch (error) {
    button.disabled = false;
    setNotice(error.message, 'error');
  }
});

for (const [input, zone] of [
  [dgrInput, document.querySelector('#dgr-dropzone')],
  [comparisonInput, document.querySelector('#comparison-dropzone')],
]) {
  for (const eventName of ['dragenter', 'dragover']) zone.addEventListener(eventName, (event) => { event.preventDefault(); zone.classList.add('drag-over'); });
  for (const eventName of ['dragleave', 'drop']) zone.addEventListener(eventName, (event) => { event.preventDefault(); zone.classList.remove('drag-over'); });
  zone.addEventListener('drop', (event) => {
    const file = event.dataTransfer.files[0];
    if (!file) return;
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    input.dispatchEvent(new Event('change', { bubbles: true }));
  });
}

document.querySelector('#run-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  if (generating) return;
  generating = true;
  updateRunButton();
  runButton.querySelector('.button-label').textContent = 'Generating...';
  setNotice('Reading workbooks and building the validation report. Larger DGR files can take a little longer.');
  const form = new FormData();
  form.append('dgr', dgrInput.files[0]);
  if (dashboardSelect.value === 'executive') form.append('comparison', comparisonInput.files[0]);
  form.append('dashboard', dashboardSelect.value);
  form.append('site', siteSelect.value);
  form.append('allDates', String(allDatesInput.checked));
  form.append('dates', JSON.stringify([...dateList.querySelectorAll('input:checked')].map((checkbox) => checkbox.value)));
  try {
    const summary = await requestJson('/api/generate', { method: 'POST', body: form });
    renderSummary(summary);
    setNotice(`Report ready for ${summary.dates.length} date${summary.dates.length === 1 ? '' : 's'}.`);
    await loadReports();
  } catch (error) {
    setNotice(error.message, 'error');
  } finally {
    generating = false;
    runButton.querySelector('.button-label').textContent = 'Generate report';
    updateRunButton();
  }
});

Promise.all([loadConfig(), loadReports()]).catch((error) => setNotice(error.message, 'error'));
// ---- Live validation (reads the app with Playwright, on the server) ----
const validateReportSelect = document.querySelector('#validate-report');
const validateSheetSelect = document.querySelector('#validate-sheet');
const validateStationInput = document.querySelector('#validate-station');
const validateDateInput = document.querySelector('#validate-date');
const validateHeadedInput = document.querySelector('#validate-headed');
const validateButton = document.querySelector('#validate-button');
const validateNotice = document.querySelector('#validate-notice');
const validateResult = document.querySelector('#validate-result');
let validating = false;

function setValidateNotice(message, kind = 'info') {
  validateNotice.textContent = message;
  validateNotice.classList.toggle('error', kind === 'error');
  validateNotice.hidden = !message;
}

function updateValidateButton() {
  validateButton.disabled = validating || !validateReportSelect.value || !validateSheetSelect.value
    || !validateStationInput.value.trim() || !validateDateInput.value;
}

async function loadValidateReports() {
  const reports = await requestJson('/api/reports');
  validateReportSelect.replaceChildren(new Option(reports.length ? 'Select a report' : 'No reports yet — generate one above', ''));
  for (const report of reports) validateReportSelect.add(new Option(report.name, report.name));
  validateReportSelect.disabled = reports.length === 0;
  validateSheetSelect.replaceChildren(new Option('Choose a report first', ''));
  validateSheetSelect.disabled = true;
  updateValidateButton();
}

validateReportSelect.addEventListener('change', async () => {
  validateSheetSelect.replaceChildren(new Option('Loading sheets...', ''));
  validateSheetSelect.disabled = true;
  updateValidateButton();
  if (!validateReportSelect.value) return;
  try {
    const { sheets } = await requestJson(`/api/reports/${encodeURIComponent(validateReportSelect.value)}/sheets`);
    validateSheetSelect.replaceChildren(new Option(sheets.length ? 'Select a sheet' : 'No sheets found', ''));
    for (const sheet of sheets) validateSheetSelect.add(new Option(sheet, sheet));
    validateSheetSelect.disabled = sheets.length === 0;
  } catch (error) {
    setValidateNotice(error.message, 'error');
  } finally {
    updateValidateButton();
  }
});

validateSheetSelect.addEventListener('change', () => {
  const match = /^(.*)_(\d{4}-\d{2}-\d{2})$/.exec(validateSheetSelect.value);
  if (match) {
    validateStationInput.value = match[1];
    validateDateInput.value = match[2];
  } else {
    validateStationInput.value = '';
    validateDateInput.value = '';
  }
  updateValidateButton();
});
[validateStationInput, validateDateInput].forEach((el) => el.addEventListener('input', updateValidateButton));

function renderValidateResult(summary) {
  validateResult.hidden = false;
  document.querySelector('#validate-result-title').textContent = `${summary.filename} · ${summary.sheet}`;
  const download = document.querySelector('#validate-download');
  download.href = `/api/reports/${encodeURIComponent(summary.filename)}`;
  download.setAttribute('download', summary.filename);
  const stats = [
    { value: summary.pass, label: 'Pass' },
    { value: summary.fail, label: 'Fail', error: summary.fail > 0 },
    { value: summary.noData, label: 'Not readable', error: summary.noData > 0 },
  ];
  document.querySelector('#validate-stats').replaceChildren(...stats.map((stat) => {
    const node = document.createElement('div');
    node.className = `stat${stat.error ? ' error' : ''}`;
    const count = document.createElement('strong');
    count.textContent = String(stat.value);
    const label = document.createElement('span');
    label.textContent = stat.label;
    node.append(count, label);
    return node;
  }));
  document.querySelector('#validate-issues').replaceChildren(...summary.issues.map((issue) => {
    const row = document.createElement('div');
    row.className = `issue-row${issue.status === 'NO_DATA' ? ' skipped' : ''}`;
    const symbol = document.createElement('span');
    symbol.className = 'issue-symbol';
    symbol.textContent = issue.status === 'NO_DATA' ? '–' : '!';
    const text = document.createElement('span');
    text.textContent = `row ${issue.row} · ${issue.section || ''} / ${issue.field || ''}: ${issue.note}`;
    row.append(symbol, text);
    return row;
  }));
  validateResult.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

document.querySelector('#validate-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  if (validating) return;
  validating = true;
  updateValidateButton();
  validateButton.querySelector('.button-label').textContent = 'Running...';
  setValidateNotice('Opening the dashboard and reading values — this can take up to a minute.');
  try {
    const summary = await requestJson('/api/validate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        report: validateReportSelect.value,
        sheet: validateSheetSelect.value,
        station: validateStationInput.value.trim(),
        date: validateDateInput.value,
        headless: !validateHeadedInput.checked,
      }),
    });
    renderValidateResult(summary);
    setValidateNotice(`Done — ${summary.pass} pass, ${summary.fail} fail, ${summary.noData} not readable.`);
    await loadReports();
  } catch (error) {
    setValidateNotice(error.message, 'error');
  } finally {
    validating = false;
    validateButton.querySelector('.button-label').textContent = 'Run validation';
    updateValidateButton();
  }
});

loadValidateReports().catch((error) => setValidateNotice(error.message, 'error'));
document.querySelector('#refresh-reports').addEventListener('click', () => loadValidateReports().catch(() => {}));
