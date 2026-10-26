'use strict';
const { chromium } = require('playwright');

/**
 * Drives the DXT dashboard like a human tester: reuse the saved login,
 * pick the station, pick the date, then read each KPI by its on-screen
 * label text.
 */
class AppReader {
  constructor({ baseUrl, storageState, headless = true }) {
    this.baseUrl = baseUrl.replace(/\/$/, '');
    this.storageState = storageState;
    this.headless = headless;
  }

  async open() {
    this.browser = await chromium.launch({ headless: this.headless });
    this.context = await this.browser.newContext({ storageState: this.storageState });
    this.page = await this.context.newPage();
    await this.page.goto(`${this.baseUrl}/dashboard`, { waitUntil: 'domcontentloaded' });
    try {
      await this.page.getByText('Select Station:', { exact: true }).waitFor({ state: 'visible', timeout: 30000 });
    } catch {
      if (await this.page.getByRole('button', { name: 'Sign In', exact: true }).isVisible().catch(() => false)) {
        throw new Error('The saved auth.json session is not signed in. Refresh it with: npx playwright codegen --save-storage=auth.json https://dxt-azure-dev.netlify.app/login');
      }
      throw new Error(`The live dashboard did not load its station selector (current URL: ${this.page.url()}).`);
    }
  }

  async close() { await this.browser.close(); }

  async screenshot(filePath) { await this.page.screenshot({ path: filePath, fullPage: true }); }

  /**
   * EDIT ME: this click path is copied from a Playwright codegen recording.
   * If the station picker on your dashboard works differently, replace the
   * body of this method with the equivalent lines from a fresh recording.
   */
  async selectStation(stationName) {
    const p = this.page;
    const stationLabel = p.getByText('Select Station:', { exact: true });
    const stationControl = stationLabel.locator('xpath=..');
    const controlText = (await stationControl.innerText()).replace(/\s+/g, ' ').trim();
    if (controlText.includes(stationName)) return;

    await stationControl.getByRole('button').click();
    const regionName = stationName.replace(/\s+\d+(?:\.\d+)?$/, '');
    const regionButton = p.getByRole('button', { name: regionName, exact: true });
    if (await regionButton.isVisible().catch(() => false)) await regionButton.click();
    const stationOption = p.getByText(stationName, { exact: true }).last();
    await stationOption.waitFor({ state: 'visible', timeout: 15000 });
    await stationOption.click();
    await p.waitForLoadState('networkidle');
  }

  async selectDate(iso) {
    const p = this.page;
    if (/^\d{2}-\d{2}-\d{4}$/.test(iso)) {
      const [day, month, year] = iso.split('-');
      iso = `${year}-${month}-${day}`;
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) throw new Error(`Invalid validation date: ${iso}`);
    const dailyMode = p.getByText('Daily', { exact: true }).first();
    await dailyMode.waitFor({ state: 'visible', timeout: 10000 });
    await dailyMode.click();
    const target = new Date(`${iso}T00:00:00Z`);
    const dayName = target.toLocaleDateString('en-US', { weekday: 'long', timeZone: 'UTC' });
    const monthName = target.toLocaleDateString('en-US', { month: 'long', timeZone: 'UTC' });
    const day = target.getUTCDate();
    const suffix = day % 100 >= 11 && day % 100 <= 13 ? 'th' : ({ 1: 'st', 2: 'nd', 3: 'rd' }[day % 10] || 'th');
    const year = target.getUTCFullYear();
    const dateButton = p.getByRole('button').filter({ hasText: /\d{2}-\d{2}-\d{4}\s+-\s+\d{2}-\d{2}-\d{4}/ }).first();
    const currentRange = (await dateButton.innerText()).replace(/\s+/g, ' ').trim();
    const currentDateMatch = /^(\d{2})-(\d{2})-(\d{4})/.exec(currentRange);
    if (!currentDateMatch) throw new Error(`Could not determine the dashboard date range: ${currentRange}`);
    const [selectedYear, month, date] = iso.split('-');
    const expectedRange = `${date}-${month}-${selectedYear} - ${date}-${month}-${selectedYear}`;
    if (currentRange === expectedRange) return;
    await dateButton.click();

    const wantedMonth = year * 12 + target.getUTCMonth();
    const currentMonth = Number(currentDateMatch[3]) * 12 + Number(currentDateMatch[2]) - 1;
    const pickerId = await dateButton.getAttribute('aria-controls');
    if (!pickerId) throw new Error('Could not identify the dashboard date picker.');
    const picker = p.locator(`[id="${pickerId}"]`);
    const dayButton = picker.locator(`button[aria-label*="${monthName} ${day}${suffix}, ${year}"]:visible`);
    const previousMonth = picker.locator('button:has(svg.lucide-chevron-left):visible').first();
    const nextMonth = picker.locator('button:has(svg.lucide-chevron-right):visible').first();
    let targetDayVisible = false;
    for (let i = 0; i < 24; i++) {
      if (await dayButton.isVisible().catch(() => false) && await dayButton.isEnabled().catch(() => false)) {
        targetDayVisible = true;
        break;
      }
      await (wantedMonth < currentMonth ? previousMonth : nextMonth).click();
    }
    if (!targetDayVisible) throw new Error(`Could not navigate the calendar to ${iso}.`);
    await dayButton.click();
    const applyButton = picker.getByRole('button', { name: 'Apply', exact: true });
    if (await applyButton.isDisabled()) await dayButton.click();
    if (await applyButton.isDisabled()) throw new Error(`The calendar did not accept ${iso} as a complete date range.`);
    await applyButton.click();
    const deadline = Date.now() + 20000;
    let selectedRange = '';
    while (Date.now() < deadline) {
      selectedRange = (await dateButton.innerText()).replace(/\s+/g, ' ').trim();
      if (selectedRange === expectedRange) break;
      await p.waitForTimeout(300);
    }
    if (selectedRange !== expectedRange) throw new Error(`Dashboard date is ${selectedRange}; expected ${expectedRange}.`);
    await p.waitForTimeout(1500);
  }

  /** Reads one KPI's on-screen text by label, e.g. readByLabel('PR') -> "71.61%". */
  async readByLabel(label, { labelRegex, chart } = {}) {
    const p = this.page;
    const matcher = labelRegex ? new RegExp(labelRegex) : label;
    const anchor = chart
      ? p.getByText(new RegExp(`^${label}`, 'i')).first()
      : p.getByText(matcher, { exact: true }).first();
    if (!(await anchor.count())) return null;
    const anchorText = (await anchor.innerText()).replace(/\s+/g, ' ').trim();
    for (let depth = 1; depth <= 3; depth++) {
      const container = anchor.locator(`xpath=ancestor::div[${depth}]`);
      const text = (await container.innerText().catch(() => '')).replace(/\s+/g, ' ').trim();
      if (!text.startsWith(anchorText)) continue;
      const value = text.slice(anchorText.length).trim();
      if (/\d|N\/A/i.test(value)) return value;
    }
    return null;
  }

  async readAll(labelConfig) {
    const out = {};
    const chartValues = await this.readEnergyGenerationValues();
    let chartValueIndex = 0;
    for (const [key, cfg] of Object.entries(labelConfig)) {
      try {
        out[key] = cfg.chart
          ? chartValues[chartValueIndex++] ?? null
          : await this.readByLabel(cfg.label, cfg);
      }
      catch { out[key] = null; }
    }
    return out;
  }

  async readEnergyGenerationValues() {
    const heading = this.page.getByRole('heading', { name: 'Energy Generation', exact: true }).first();
    if (!(await heading.count())) return [];
    const card = heading.locator('xpath=ancestor::div[contains(@class, "rounded-md")][1]');
    const values = card.locator('div[class*="text-[10px]"][class*="font-bold"]:visible');
    return (await values.allTextContents()).map((value) => value.trim());
  }
}

module.exports = { AppReader };
