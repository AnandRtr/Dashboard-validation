'use strict';

const fs = require('node:fs');
const { request } = require('playwright');

const SUMMARY_PATH = '/api/v1/metrics/plant/kpis/summary';

function readAccessToken(authPath) {
  const storageState = JSON.parse(fs.readFileSync(authPath, 'utf8'));
  const tokenEntry = (storageState.origins || [])
    .flatMap((origin) => origin.localStorage || [])
    .find((entry) => entry.name === 'accessToken');
  const token = tokenEntry && tokenEntry.value;
  if (typeof token !== 'string' || !token.trim()) {
    throw new Error('No accessToken found in auth.json local storage. Refresh the saved session.');
  }
  return /^Bearer\s/i.test(token) ? token : `Bearer ${token}`;
}

class LiveClient {
  constructor({ baseApiUrl, authPath }) {
    this.baseApiUrl = baseApiUrl.replace(/\/$/, '');
    this.authPath = authPath;
  }

  async fetchDaySummary(stationId, appDate) {
    const authorization = readAccessToken(this.authPath);
    const api = await request.newContext({
      baseURL: this.baseApiUrl,
      storageState: this.authPath,
      extraHTTPHeaders: { Authorization: authorization },
    });

    try {
      const response = await api.get(SUMMARY_PATH, {
        params: { stationId, startDate: appDate, endDate: appDate },
      });
      if (response.status() === 401 || response.status() === 403) {
        throw new Error('The saved auth.json session is not authorized for the live API. Refresh the saved session.');
      }
      if (!response.ok()) throw new Error(`Live API request failed with status ${response.status()}.`);

      let payload;
      try {
        payload = await response.json();
      } catch {
        throw new Error('The live API returned an invalid JSON response.');
      }
      if (payload.status !== 'success' || !payload.data || typeof payload.data !== 'object') {
        throw new Error('The live API response did not contain a successful day summary.');
      }
      return payload.data;
    } finally {
      await api.dispose();
    }
  }
}

function getByPath(value, dotPath) {
  if (typeof dotPath !== 'string' || !dotPath) return undefined;
  return dotPath.split('.').reduce((current, key) => current == null ? undefined : current[key], value);
}

module.exports = { LiveClient, getByPath };