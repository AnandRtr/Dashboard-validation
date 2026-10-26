import { test, expect } from '@playwright/test';

test.use({ storageState: 'auth.json' });

test('dashboard values for the selected station and date', async ({ page }) => {
  await page.goto('https://dxt-azure-dev.netlify.app/dashboard');
  await page.getByRole('button', { name: 'Andhra Pradesh' }).click();
  await page.locator('div').filter({ hasText: /^Andhra Pradesh 2\.1South • Andhra Pradesh • IoT$/ }).first().click();
  await page.getByRole('button').filter({ hasText: /^$/ }).nth(1).click();
  await page.getByRole('button', { name: 'Go to the Previous Month' }).click();
  await page.getByRole('button', { name: 'Monday, August 10th,' }).click();
  await page.locator('div').filter({ hasText: /^Energy Export287\.64 MWhSource: FALLBACKUpdated: 11-08-2026 05:29:00$/ }).nth(1).click();
  await page.locator('div').filter({ hasText: /^Energy Import0\.00 MWhSource: NOT_FOUNDUpdated: 11-08-2026 05:29:00$/ }).nth(1).click();
  await page.locator('div').filter({ hasText: /^Net Generation287\.64 MWh$/ }).first().click();
  await page.locator('div').filter({ hasText: /^Total Inv Generation287\.64 MWh$/ }).nth(1).click();
  await page.locator('div').filter({ hasText: /^Active PowerN\/A$/ }).nth(1).click();
  await page.locator('div').filter({ hasText: /^Peak Power50\.02 MWUpdated: 10-08-2026 11:56:00$/ }).nth(1).click();
  await page.locator('div').filter({ hasText: /^PR71\.61%$/ }).nth(1).click();
  await page.locator('div').filter({ hasText: /^PLF18\.26%$/ }).first().click();
  await expect(page.getByRole('main')).toMatchAriaSnapshot(`
    - img
    - text: /GHI \\d+\\.\\d+ kWh\\/m²/
    `);
  await page.getByText('GTI6.12 kWh/m²').click();
});