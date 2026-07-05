#!/usr/bin/env python3
import json
import sqlite3
import uuid
from datetime import datetime

DB_PATH = "/Users/v/Documents/N8N_LOCAL/.n8n/database.sqlite"
WORKFLOW_ID = "tQLIIUtCOFIcRbMo"


def now_sql():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]


def node_id():
    return str(uuid.uuid4())


PARSE_CONFIG_JS = r"""
const crypto = require('crypto');

function truthy(value) {
  return String(value ?? '').toLowerCase() === 'true';
}

function parseLegacyCredentials(text) {
  const creds = {};
  for (const rawLine of String(text || '').split('\n')) {
    let line = rawLine.trim();
    if (!line) continue;
    if (line.startsWith('=')) line = line.slice(1);
    const parts = line.split('\t');
    if (parts.length < 2) continue;
    let key = parts[0].trim().replace(/:$/, '');
    let value = parts.slice(1).join('\t');
    if (value.includes('\\n')) value = value.replace(/\\n/g, '\n');
    creds[key] = value;
  }
  return creds;
}

function base64url(input) {
  return Buffer.from(input).toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
}

const raw = $input.first().json;
const legacy = parseLegacyCredentials(raw.LEGACY_CREDENTIALS);

let maxRows = String(raw.MAX_ROWS_PER_RUN || 'ALL').trim();
const legacyMaxRows =
  legacy['✅GOOGLE_SHEET_MIDJOURNEY_APIKEY_AMMOUNT_OF_LINES_TO_READ'] ||
  legacy['✅GOOGLE_SHEET_MIDJOURNEY_APIKEY_midapi.ai_AMMOUNT_OF_LINES_TO_READ'];
if (maxRows.toUpperCase() === 'ALL' && legacyMaxRows) maxRows = String(legacyMaxRows).trim();

const cfg = {
  TTAPI_API_KEY: raw.TTAPI_API_KEY || legacy['✅MIDJOURNEY_APIKEY'] || legacy.TTAPI_API_KEY,
  GOOGLE_SHEET_ID: raw.GOOGLE_SHEET_ID || legacy['✅GOOGLE_SHEET_ID'],
  GOOGLE_SERVICE_ACCOUNT_EMAIL: raw.GOOGLE_SERVICE_ACCOUNT_EMAIL || legacy['✅GOOGLE_SERVICE_ACCOUNT_EMAIL'],
  GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY: raw.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || legacy['✅GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY'],
  TELEGRAM_BOT_TOKEN: raw.TELEGRAM_BOT_TOKEN || legacy['✅TELEGRAM_BOT_API'],
  TELEGRAM_CHAT_ID: raw.TELEGRAM_CHAT_ID || legacy['✅TELEGRAM_MESSAGE_TO'],
  SHEET_NAME: raw.SHEET_NAME || 'MIDJOURNEY_midapi.ai',
  MODE: raw.MODE || 'relax',
  TIMEOUT_SECONDS: Number(raw.TIMEOUT_SECONDS || 1200),
  MAX_IN_FLIGHT: Math.max(1, Math.min(10, Number(raw.MAX_IN_FLIGHT || 5))),
  POLL_SECONDS: Math.max(10, Number(raw.POLL_SECONDS || 30)),
  MAX_RUNTIME_MINUTES: Math.max(1, Number(raw.MAX_RUNTIME_MINUTES || 55)),
  MAX_ROWS_PER_RUN: maxRows || 'ALL',
  SEND_TELEGRAM: truthy(raw.SEND_TELEGRAM),
  WRITE_IMAGE_FORMULAS: truthy(raw.WRITE_IMAGE_FORMULAS),
  AUTO_FIX_SINGLE_DASH_OW: truthy(raw.AUTO_FIX_SINGLE_DASH_OW),
  OUTPUT_DIR: raw.OUTPUT_DIR || '/Users/v/Documents/ART_GPT/ttapi_downloads',
};

for (const key of ['TTAPI_API_KEY', 'GOOGLE_SHEET_ID', 'GOOGLE_SERVICE_ACCOUNT_EMAIL', 'GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY']) {
  if (!cfg[key]) throw new Error(`Missing required config: ${key}`);
}

const now = Math.floor(Date.now() / 1000);
const header = { alg: 'RS256', typ: 'JWT' };
const claim = {
  iss: cfg.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  scope: 'https://www.googleapis.com/auth/spreadsheets',
  aud: 'https://oauth2.googleapis.com/token',
  exp: now + 3600,
  iat: now,
};
const unsigned = base64url(JSON.stringify(header)) + '.' + base64url(JSON.stringify(claim));
const signature = crypto.createSign('RSA-SHA256').update(unsigned).sign(cfg.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY, 'base64')
  .replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');

cfg.jwtAssertion = unsigned + '.' + signature;
cfg.runStartedAt = new Date().toISOString();
cfg.SHEET_A1_PREFIX = `'${String(cfg.SHEET_NAME).replace(/'/g, "''")}'!`;

return [{ json: cfg }];
"""


SELECT_ROWS_JS = r"""
const cfg = $('Parse Config & JWT').first().json;
const googleToken = $('Get Google Token').first().json.access_token;
const rows = $input.first().json.values || [];

function statusOf(row) {
  return String(row[5] || '').trim().toUpperCase();
}

function doneCell(value) {
  const text = String(value || '').trim();
  return text && !/^ERROR$/i.test(text) && !/^FAILED/i.test(text);
}

function normalizePrompt(prompt) {
  let output = String(prompt || '').trim();
  const warnings = [];
  const errors = [];
  if (!output) errors.push('Prompt is empty');

  if (cfg.AUTO_FIX_SINGLE_DASH_OW) {
    const fixed = output.replace(/(^|\s)-ow(\s+\d+)/g, '$1--ow$2');
    if (fixed !== output) {
      output = fixed;
      warnings.push('Fixed single-dash -ow to --ow');
    }
  }

  for (const match of output.matchAll(/--iw\s+([0-9.]+)/g)) {
    const value = Number(match[1]);
    if (Number.isFinite(value) && value > 3) {
      errors.push(`Invalid --iw ${match[1]}: Midjourney image weight must be 0-3. Use --ow for Omni Reference weight.`);
    }
  }

  if (/https:\/\/ibb\.co\//i.test(output)) {
    warnings.push('Prompt contains ibb.co page URL; direct i.ibb.co image URL is safer for API generation.');
  }

  return { prompt: output, warnings, errors };
}

const maxNewRows = String(cfg.MAX_ROWS_PER_RUN || 'ALL').toUpperCase() === 'ALL'
  ? Infinity
  : Math.max(0, Number(cfg.MAX_ROWS_PER_RUN) || 0);

const running = [];
const ready = [];

for (let i = 1; i < rows.length; i++) {
  const rowNumber = i + 1;
  const row = rows[i] || [];
  const prompt = String(row[0] || '').trim();
  if (!prompt) continue;

  const status = statusOf(row);
  const jobId = String(row[6] || '').trim();
  const image1 = row[1];
  const attempts = Number(row[11] || 0) || 0;

  if ((status === 'SUBMITTED' || status === 'RUNNING') && jobId) {
    running.push({ rowNumber, prompt, jobId, attempts, submittedAt: row[9] || '', resume: true });
    continue;
  }

  if (status === 'DONE' || doneCell(image1) || status === 'LOCKED') continue;
  if (!status || ['READY', 'FAILED', 'ERROR', 'FAILED_PROMPT', 'RETRY'].includes(status) || /^ERROR$/i.test(String(image1 || ''))) {
    ready.push({ rowNumber, prompt, attempts, resume: false });
  }
}

const selected = [];
for (const item of running.slice(0, cfg.MAX_IN_FLIGHT)) selected.push(item);

let newCount = 0;
while (selected.length < cfg.MAX_IN_FLIGHT && ready.length && newCount < maxNewRows) {
  selected.push(ready.shift());
  newCount++;
}

if (selected.length === 0) {
  return [{ json: { noWork: true, message: 'No READY/SUBMITTED/RUNNING rows found.' } }];
}

return selected.map(item => {
  const normalized = normalizePrompt(item.prompt);
  const attempts = item.resume ? item.attempts : item.attempts + 1;
  return {
    json: {
      ...cfg,
      googleToken,
      noWork: false,
      rowNumber: item.rowNumber,
      originalPrompt: item.prompt,
      prompt: normalized.prompt,
      warnings: normalized.warnings,
      errors: normalized.errors,
      promptValid: normalized.errors.length === 0,
      hasJob: Boolean(item.jobId),
      jobId: item.jobId || '',
      attempts,
      submittedAt: item.submittedAt || new Date().toISOString(),
      deadlineMs: Date.now() + cfg.TIMEOUT_SECONDS * 1000,
    }
  };
});
"""


SHEET_UPDATE_HELPER_JS = r"""
const https = require('https');
const { URL } = require('url');

function requestJson(method, url, headers, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const payload = body === undefined ? undefined : JSON.stringify(body);
    const req = https.request({
      method,
      hostname: u.hostname,
      path: u.pathname + u.search,
      headers: {
        ...headers,
        ...(payload ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) } : {}),
      },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        let json = {};
        try { json = data ? JSON.parse(data) : {}; } catch { json = { raw: data }; }
        if (res.statusCode < 200 || res.statusCode >= 300) reject(new Error(JSON.stringify(json)));
        else resolve(json);
      });
    });
    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

function sheetRange(item, range) {
  const prefix = item.SHEET_A1_PREFIX || `'${String(item.SHEET_NAME).replace(/'/g, "''")}'!`;
  return encodeURIComponent(`${prefix}${range}`);
}

async function updateRow(item, values) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${item.GOOGLE_SHEET_ID}/values/${sheetRange(item, `B${item.rowNumber}:M${item.rowNumber}`)}?valueInputOption=USER_ENTERED`;
  return requestJson('PUT', url, { Authorization: `Bearer ${item.googleToken}` }, { values: [values] });
}
"""


MARK_PROMPT_FAILED_JS = SHEET_UPDATE_HELPER_JS + r"""
const out = [];
for (const { json: item } of $input.all()) {
  await updateRow(item, [
    '', '', '', '',
    'FAILED_PROMPT',
    item.jobId || '',
    '',
    (item.errors || []).join('; ').slice(0, 45000),
    item.submittedAt || '',
    new Date().toISOString(),
    item.attempts || '',
    '',
  ]);
  out.push({ json: { ...item, pollState: 'FINAL_FAILED', message: (item.errors || []).join('; ') } });
}
return out;
"""


LOCK_ROW_JS = SHEET_UPDATE_HELPER_JS + r"""
const out = [];
for (const { json: item } of $input.all()) {
  const message = (item.warnings || []).length ? item.warnings.join('; ') : 'Preparing TTAPI submit';
  await updateRow(item, [
    '', '', '', '',
    'LOCKED',
    '',
    '',
    message,
    item.submittedAt,
    '',
    item.attempts,
    '',
  ]);
  out.push({ json: item });
}
return out;
"""


SUBMIT_TTAPI_JS = SHEET_UPDATE_HELPER_JS + r"""
function ttapi(method, url, key, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const payload = body === undefined ? undefined : JSON.stringify(body);
    const req = https.request({
      method,
      hostname: u.hostname,
      path: u.pathname + u.search,
      headers: {
        'TT-API-KEY': key,
        'User-Agent': 'n8n-ttapi-v1/1.0',
        ...(payload ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) } : {}),
      },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        let json = {};
        try { json = data ? JSON.parse(data) : {}; } catch { json = { raw: data }; }
        if (res.statusCode < 200 || res.statusCode >= 300) reject(new Error(json.message || JSON.stringify(json)));
        else resolve(json);
      });
    });
    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

const out = [];
for (const { json: item } of $input.all()) {
  try {
    const res = await ttapi('POST', 'https://api.ttapi.io/midjourney/v1/imagine', item.TTAPI_API_KEY, {
      prompt: item.prompt,
      mode: item.MODE,
      timeout: item.TIMEOUT_SECONDS,
    });
    if (res.status === 'FAILED' || !res.data?.jobId) throw new Error(res.message || JSON.stringify(res));
    const jobId = res.data.jobId;
    await updateRow(item, [
      '', '', '', '',
      'SUBMITTED',
      jobId,
      '',
      (item.warnings || []).length ? item.warnings.join('; ') : 'Submitted',
      item.submittedAt,
      '',
      item.attempts,
      '',
    ]);
    out.push({ json: { ...item, jobId, hasJob: true, pollState: 'RUNNING' } });
  } catch (error) {
    await updateRow(item, [
      '', '', '', '',
      'FAILED',
      '',
      '',
      String(error.message || error).slice(0, 45000),
      item.submittedAt,
      new Date().toISOString(),
      item.attempts,
      '',
    ]);
    out.push({ json: { ...item, pollState: 'FINAL_FAILED', message: error.message } });
  }
}
return out;
"""


POLL_TTAPI_JS = SHEET_UPDATE_HELPER_JS + r"""
function ttapiFetch(jobId, key) {
  return new Promise((resolve, reject) => {
    const u = new URL(`https://api.ttapi.io/midjourney/v1/fetch?jobId=${encodeURIComponent(jobId)}`);
    const req = https.request({
      method: 'GET',
      hostname: u.hostname,
      path: u.pathname + u.search,
      headers: { 'TT-API-KEY': key, 'User-Agent': 'n8n-ttapi-v1/1.0' },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        let json = {};
        try { json = data ? JSON.parse(data) : {}; } catch { json = { raw: data }; }
        if (res.statusCode < 200 || res.statusCode >= 300) reject(new Error(json.message || JSON.stringify(json)));
        else resolve(json);
      });
    });
    req.on('error', reject);
    req.end();
  });
}

const out = [];
for (const { json: item } of $input.all()) {
  if (!item.jobId) {
    out.push({ json: { ...item, pollState: 'FINAL_FAILED', message: 'Missing jobId' } });
    continue;
  }

  if (Date.now() > item.deadlineMs) {
    await updateRow(item, [
      '', '', '', '',
      'FAILED_TIMEOUT',
      item.jobId,
      '',
      'Exceeded local polling timeout',
      item.submittedAt,
      new Date().toISOString(),
      item.attempts,
      '',
    ]);
    out.push({ json: { ...item, pollState: 'FINAL_FAILED', message: 'Exceeded local polling timeout' } });
    continue;
  }

  try {
    const res = await ttapiFetch(item.jobId, item.TTAPI_API_KEY);
    const data = res.data || {};
    if (res.status === 'SUCCESS') {
      out.push({
        json: {
          ...item,
          pollState: 'DONE',
          gridImage: data.cdnImage || data.discordImage || '',
          imageUrls: (data.images || []).filter(Boolean),
          resultPrompt: data.prompt || item.prompt,
        }
      });
    } else if (res.status === 'FAILED') {
      const message = res.message || data.error || 'TTAPI failed';
      await updateRow(item, [
        '', '', '', '',
        'FAILED',
        item.jobId,
        '',
        String(message).slice(0, 45000),
        item.submittedAt,
        new Date().toISOString(),
        item.attempts,
        '',
      ]);
      out.push({ json: { ...item, pollState: 'FINAL_FAILED', message } });
    } else {
      const progress = data.progress ? `Progress ${data.progress}%` : (res.status || 'RUNNING');
      await updateRow(item, [
        '', '', '', '',
        'RUNNING',
        item.jobId,
        '',
        progress,
        item.submittedAt,
        '',
        item.attempts,
        '',
      ]);
      out.push({ json: { ...item, pollState: 'RUNNING', message: progress } });
    }
  } catch (error) {
    await updateRow(item, [
      '', '', '', '',
      'RUNNING',
      item.jobId,
      '',
      `Poll error; will retry: ${String(error.message || error).slice(0, 300)}`,
      item.submittedAt,
      '',
      item.attempts,
      '',
    ]);
    out.push({ json: { ...item, pollState: 'RUNNING', message: error.message } });
  }
}
return out;
"""


FINALIZE_JS = SHEET_UPDATE_HELPER_JS + r"""
const fs = require('fs');
const path = require('path');

function safeName(value) {
  return String(value || '').replace(/[^a-zA-Z0-9_.-]+/g, '_').slice(0, 100);
}

function asImageCell(item, url) {
  if (!url) return '';
  const escaped = String(url).replace(/"/g, '""');
  return item.WRITE_IMAGE_FORMULAS ? `=IMAGE("${escaped}")` : url;
}

function download(url, file) {
  return new Promise((resolve, reject) => {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    const u = new URL(url);
    const req = https.request({ method: 'GET', hostname: u.hostname, path: u.pathname + u.search }, res => {
      if (res.statusCode < 200 || res.statusCode >= 300) {
        reject(new Error(`Download HTTP ${res.statusCode}`));
        return;
      }
      const stream = fs.createWriteStream(file);
      res.pipe(stream);
      stream.on('finish', () => stream.close(() => resolve(file)));
      stream.on('error', reject);
    });
    req.on('error', reject);
    req.end();
  });
}

function sendTelegram(item, url, caption) {
  if (!item.SEND_TELEGRAM || !item.TELEGRAM_BOT_TOKEN || !item.TELEGRAM_CHAT_ID || !url) {
    return Promise.resolve({ skipped: true });
  }
  return requestJson('POST', `https://api.telegram.org/bot${item.TELEGRAM_BOT_TOKEN}/sendDocument`, {}, {
    chat_id: item.TELEGRAM_CHAT_ID,
    document: url,
    caption,
  });
}

const out = [];
for (const { json: item } of $input.all()) {
  if (item.pollState !== 'DONE') {
    out.push({ json: { rowNumber: item.rowNumber, jobId: item.jobId, status: item.pollState, message: item.message || '' } });
    continue;
  }

  const urls = item.imageUrls || [];
  const localFiles = [];
  let telegramOk = 0;

  for (let i = 0; i < urls.length; i++) {
    const ext = path.extname(new URL(urls[i]).pathname) || '.png';
    const file = path.join(item.OUTPUT_DIR, `row_${item.rowNumber}_job_${safeName(item.jobId)}_${i + 1}${ext}`);
    try {
      await download(urls[i], file);
      localFiles.push(file);
    } catch (error) {
      localFiles.push(`DOWNLOAD_ERROR: ${error.message}`);
    }

    try {
      await sendTelegram(item, urls[i], `Row ${item.rowNumber} image ${i + 1}/${urls.length}`);
      telegramOk += 1;
    } catch (error) {
      localFiles.push(`TELEGRAM_ERROR: ${error.message}`);
    }
  }

  const finalStatus = telegramOk === urls.length ? 'DONE' : 'DONE_TELEGRAM_ERROR';
  await updateRow(item, [
    asImageCell(item, urls[0]),
    asImageCell(item, urls[1]),
    asImageCell(item, urls[2]),
    asImageCell(item, urls[3]),
    finalStatus,
    item.jobId,
    item.gridImage || '',
    localFiles.join('\n').slice(0, 45000),
    item.submittedAt,
    new Date().toISOString(),
    item.attempts,
    `${telegramOk}/${urls.length}`,
  ]);

  out.push({ json: { rowNumber: item.rowNumber, jobId: item.jobId, status: finalStatus, images: urls.length, telegramSent: `${telegramOk}/${urls.length}`, localFiles } });
}
return out;
"""


def make_http_node(name, method, url, position, headers=None, body=None, content_type=None):
    params = {
        "method": method,
        "url": url,
        "sendHeaders": bool(headers),
        "options": {},
    }
    if headers:
        params["headerParameters"] = {"parameters": headers}
    if body is not None:
        params["sendBody"] = True
        if content_type == "form-urlencoded":
            params["contentType"] = "form-urlencoded"
            params["bodyParameters"] = {"parameters": body}
        else:
            params["specifyBody"] = "json"
            params["jsonBody"] = body
    return {
        "parameters": params,
        "type": "n8n-nodes-base.httpRequest",
        "typeVersion": 4.2,
        "position": position,
        "id": node_id(),
        "name": name,
    }


def make_code_node(name, js, position):
    return {
        "parameters": {"mode": "runOnceForAllItems", "jsCode": js},
        "type": "n8n-nodes-base.code",
        "typeVersion": 2,
        "position": position,
        "id": node_id(),
        "name": name,
    }


def make_if_node(name, left, right, op_type, operation, position):
    return {
        "parameters": {
            "conditions": {
                "options": {"caseSensitive": True, "leftValue": "", "typeValidation": "strict", "version": 2},
                "conditions": [{
                    "id": node_id(),
                    "leftValue": left,
                    "rightValue": right,
                    "operator": {"type": op_type, "operation": operation},
                }],
                "combinator": "and",
            },
            "options": {},
        },
        "type": "n8n-nodes-base.if",
        "typeVersion": 2.2,
        "position": position,
        "id": node_id(),
        "name": name,
    }


def main():
    conn = sqlite3.connect(DB_PATH)
    try:
        row = conn.execute("SELECT nodes FROM workflow_entity WHERE id=?", (WORKFLOW_ID,)).fetchone()
        if not row:
            raise SystemExit(f"Workflow not found: {WORKFLOW_ID}")
        current_nodes = json.loads(row[0])
        config_node = next((n for n in current_nodes if n.get("name") == "Edit Config"), None)
        if not config_node:
            raise SystemExit("Edit Config node not found")

        # Keep the user's current config edits, but make test runs conservative.
        assignments = config_node["parameters"]["assignments"]["assignments"]
        for assignment in assignments:
            if assignment.get("name") == "MAX_ROWS_PER_RUN" and str(assignment.get("value", "")).upper() == "ALL":
                assignment["value"] = "1"

        nodes = [
            {
                "parameters": {
                    "content": (
                        "## ttapi_v1 visible flow\n"
                        "No TTAPI credits are spent until the `Submit TTAPI Job` node runs.\n\n"
                        "This workflow reads one prompt row by default, validates it, locks the row, submits TTAPI Relax, polls until done, downloads files, sends Telegram, and updates the sheet.\n\n"
                        "For larger batches, raise `MAX_ROWS_PER_RUN` and `MAX_IN_FLIGHT` in `Edit Config`."
                    ),
                    "height": 300,
                    "width": 460,
                    "color": 4,
                },
                "type": "n8n-nodes-base.stickyNote",
                "typeVersion": 1,
                "position": [-220, -320],
                "id": node_id(),
                "name": "README",
            },
            {
                "parameters": {},
                "type": "n8n-nodes-base.manualTrigger",
                "typeVersion": 1,
                "position": [-220, 80],
                "id": node_id(),
                "name": "Manual Start",
            },
            {
                "parameters": {"httpMethod": "POST", "path": "ttapi-v1-run", "options": {}},
                "type": "n8n-nodes-base.webhook",
                "typeVersion": 1.1,
                "position": [-220, 260],
                "id": node_id(),
                "name": "Webhook Start",
                "webhookId": str(uuid.uuid4()),
            },
            {**config_node, "position": [40, 170]},
            make_code_node("Parse Config & JWT", PARSE_CONFIG_JS, [300, 170]),
            make_http_node(
                "Get Google Token",
                "POST",
                "https://oauth2.googleapis.com/token",
                [560, 170],
                headers=[{"name": "Content-Type", "value": "application/x-www-form-urlencoded"}],
                content_type="form-urlencoded",
                body=[
                    {"name": "grant_type", "value": "urn:ietf:params:oauth:grant-type:jwt-bearer"},
                    {"name": "assertion", "value": "={{ $json.jwtAssertion }}"},
                ],
            ),
            make_http_node(
                "Ensure Sheet Headers",
                "PUT",
                "={{ 'https://sheets.googleapis.com/v4/spreadsheets/' + $('Parse Config & JWT').first().json.GOOGLE_SHEET_ID + '/values/' + encodeURIComponent($('Parse Config & JWT').first().json.SHEET_A1_PREFIX + 'F1:M1') + '?valueInputOption=USER_ENTERED' }}",
                [820, 170],
                headers=[
                    {"name": "Authorization", "value": "=Bearer {{ $('Get Google Token').first().json.access_token }}"},
                    {"name": "Content-Type", "value": "application/json"},
                ],
                body='={{ JSON.stringify({ values: [[ "STATUS", "TTAPI_JOB_ID", "GRID_IMAGE_URL", "ERROR_OR_PROGRESS", "SUBMITTED_AT", "COMPLETED_AT", "ATTEMPTS", "TELEGRAM_SENT" ]] }) }}',
            ),
            make_http_node(
                "Read Sheet",
                "GET",
                "={{ 'https://sheets.googleapis.com/v4/spreadsheets/' + $('Parse Config & JWT').first().json.GOOGLE_SHEET_ID + '/values/' + encodeURIComponent($('Parse Config & JWT').first().json.SHEET_A1_PREFIX + 'A:M') }}",
                [1080, 170],
                headers=[{"name": "Authorization", "value": "=Bearer {{ $('Get Google Token').first().json.access_token }}"}],
            ),
            make_code_node("Select Rows & Validate Prompts", SELECT_ROWS_JS, [1340, 170]),
            make_if_node("Has Work?", "={{ $json.noWork }}", False, "boolean", "equals", [1580, 170]),
            make_if_node("Prompt Valid?", "={{ $json.promptValid }}", True, "boolean", "equals", [1820, 90]),
            make_code_node("Mark Prompt Failed", MARK_PROMPT_FAILED_JS, [2060, 260]),
            make_if_node("Already Submitted?", "={{ $json.hasJob }}", True, "boolean", "equals", [2060, 90]),
            make_code_node("Lock Row", LOCK_ROW_JS, [2300, 180]),
            make_code_node("Submit TTAPI Job", SUBMIT_TTAPI_JS, [2540, 180]),
            {
                "parameters": {"amount": "={{ $('Parse Config & JWT').first().json.POLL_SECONDS }}", "unit": "seconds"},
                "type": "n8n-nodes-base.wait",
                "typeVersion": 1.1,
                "position": [2780, 90],
                "id": node_id(),
                "name": "Wait Before Poll",
                "webhookId": str(uuid.uuid4()),
            },
            make_code_node("Poll TTAPI Status", POLL_TTAPI_JS, [3020, 90]),
            make_if_node("Still Running?", "={{ $json.pollState }}", "RUNNING", "string", "equals", [3260, 90]),
            make_code_node("Finalize Results", FINALIZE_JS, [3500, 220]),
            make_code_node("Execution Summary", "return $input.all();", [3740, 220]),
        ]

        connections = {
            "Manual Start": {"main": [[{"node": "Edit Config", "type": "main", "index": 0}]]},
            "Webhook Start": {"main": [[{"node": "Edit Config", "type": "main", "index": 0}]]},
            "Edit Config": {"main": [[{"node": "Parse Config & JWT", "type": "main", "index": 0}]]},
            "Parse Config & JWT": {"main": [[{"node": "Get Google Token", "type": "main", "index": 0}]]},
            "Get Google Token": {"main": [[{"node": "Ensure Sheet Headers", "type": "main", "index": 0}]]},
            "Ensure Sheet Headers": {"main": [[{"node": "Read Sheet", "type": "main", "index": 0}]]},
            "Read Sheet": {"main": [[{"node": "Select Rows & Validate Prompts", "type": "main", "index": 0}]]},
            "Select Rows & Validate Prompts": {"main": [[{"node": "Has Work?", "type": "main", "index": 0}]]},
            "Has Work?": {"main": [[{"node": "Prompt Valid?", "type": "main", "index": 0}], []]},
            "Prompt Valid?": {"main": [[{"node": "Already Submitted?", "type": "main", "index": 0}], [{"node": "Mark Prompt Failed", "type": "main", "index": 0}]]},
            "Mark Prompt Failed": {"main": [[{"node": "Execution Summary", "type": "main", "index": 0}]]},
            "Already Submitted?": {"main": [[{"node": "Wait Before Poll", "type": "main", "index": 0}], [{"node": "Lock Row", "type": "main", "index": 0}]]},
            "Lock Row": {"main": [[{"node": "Submit TTAPI Job", "type": "main", "index": 0}]]},
            "Submit TTAPI Job": {"main": [[{"node": "Wait Before Poll", "type": "main", "index": 0}]]},
            "Wait Before Poll": {"main": [[{"node": "Poll TTAPI Status", "type": "main", "index": 0}]]},
            "Poll TTAPI Status": {"main": [[{"node": "Still Running?", "type": "main", "index": 0}]]},
            "Still Running?": {"main": [[{"node": "Wait Before Poll", "type": "main", "index": 0}], [{"node": "Finalize Results", "type": "main", "index": 0}]]},
            "Finalize Results": {"main": [[{"node": "Execution Summary", "type": "main", "index": 0}]]},
        }

        conn.execute(
            """
            UPDATE workflow_entity
            SET nodes=?, connections=?, updatedAt=?, versionId=?, versionCounter=versionCounter+1
            WHERE id=?
            """,
            (
                json.dumps(nodes, separators=(",", ":")),
                json.dumps(connections, separators=(",", ":")),
                now_sql(),
                str(uuid.uuid4()),
                WORKFLOW_ID,
            ),
        )
        conn.commit()
        print(json.dumps({"workflowId": WORKFLOW_ID, "nodes": len(nodes), "name": "ttapi_v1"}))
    finally:
        conn.close()


if __name__ == "__main__":
    main()
