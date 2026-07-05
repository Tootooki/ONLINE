#!/usr/bin/env python3
import json
import secrets
import sqlite3
import string
import uuid
from datetime import datetime

DB_PATH = "/Users/v/Documents/N8N_LOCAL/.n8n/database.sqlite"
SOURCE_WORKFLOW_ID = "UbOUgmPAK2L7KJxY"
NEW_WORKFLOW_NAME = "ttapi_v1"
SHEET_NAME = "MIDJOURNEY_midapi.ai"
OUTPUT_DIR = "/Users/v/Documents/ART_GPT/ttapi_downloads"


def n8n_id(length=16):
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(length))


def now_sql():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]


WORKER_JS = r"""
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const cfg = $input.first().json;

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
    let key = parts[0].trim();
    if (key.startsWith('=')) key = key.slice(1);
    key = key.replace(/:$/, '');
    let value = parts.slice(1).join('\t');
    if (value.includes('\\n')) value = value.replace(/\\n/g, '\n');
    creds[key] = value;
  }
  return creds;
}

const legacy = parseLegacyCredentials(cfg.LEGACY_CREDENTIALS);
const TTAPI_KEY = cfg.TTAPI_API_KEY || legacy.TTAPI_API_KEY;
const TELEGRAM_BOT_TOKEN = cfg.TELEGRAM_BOT_TOKEN || legacy['✅TELEGRAM_BOT_API'];
const TELEGRAM_CHAT_ID = cfg.TELEGRAM_CHAT_ID || legacy['✅TELEGRAM_MESSAGE_TO'];
const GOOGLE_SHEET_ID = cfg.GOOGLE_SHEET_ID || legacy['✅GOOGLE_SHEET_ID'];
const GOOGLE_SERVICE_ACCOUNT_EMAIL = cfg.GOOGLE_SERVICE_ACCOUNT_EMAIL || legacy['✅GOOGLE_SERVICE_ACCOUNT_EMAIL'];
const GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY = cfg.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || legacy['✅GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY'];
const SHEET_NAME = cfg.SHEET_NAME || 'MIDJOURNEY_midapi.ai';
const MODE = cfg.MODE || 'relax';
const TIMEOUT_SECONDS = Number(cfg.TIMEOUT_SECONDS || 1200);
const MAX_IN_FLIGHT = Math.max(1, Math.min(10, Number(cfg.MAX_IN_FLIGHT || 5)));
const POLL_SECONDS = Math.max(10, Number(cfg.POLL_SECONDS || 30));
const MAX_RUNTIME_MINUTES = Math.max(1, Number(cfg.MAX_RUNTIME_MINUTES || 55));
const MAX_ROWS_PER_RUN = String(cfg.MAX_ROWS_PER_RUN || 'ALL').trim();
const SEND_TELEGRAM = truthy(cfg.SEND_TELEGRAM);
const WRITE_IMAGE_FORMULAS = truthy(cfg.WRITE_IMAGE_FORMULAS);
const AUTO_FIX_SINGLE_DASH_OW = truthy(cfg.AUTO_FIX_SINGLE_DASH_OW);
const OUTPUT_DIR = cfg.OUTPUT_DIR || '/Users/v/Documents/ART_GPT/ttapi_downloads';

for (const [name, value] of Object.entries({
  TTAPI_KEY,
  GOOGLE_SHEET_ID,
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY,
})) {
  if (!value) throw new Error(`Missing required config: ${name}`);
}

function base64url(input) {
  return Buffer.from(input)
    .toString('base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
}

async function sleep(ms) {
  await new Promise(resolve => setTimeout(resolve, ms));
}

function isoNow() {
  return new Date().toISOString();
}

function sheetRange(range) {
  return encodeURIComponent(`${SHEET_NAME}!${range}`);
}

async function getGoogleToken() {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const claim = {
    iss: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
  };
  const unsigned = `${base64url(JSON.stringify(header))}.${base64url(JSON.stringify(claim))}`;
  const signature = crypto
    .createSign('RSA-SHA256')
    .update(unsigned)
    .sign(GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY, 'base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');

  const form = new URLSearchParams();
  form.set('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
  form.set('assertion', `${unsigned}.${signature}`);

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: form,
  });
  const json = await res.json();
  if (!res.ok || !json.access_token) {
    throw new Error(`Google token failed: ${JSON.stringify(json)}`);
  }
  return json.access_token;
}

let googleToken = await getGoogleToken();

async function sheets(method, range, body) {
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}/values/${sheetRange(range)}?valueInputOption=USER_ENTERED`,
    {
      method,
      headers: {
        Authorization: `Bearer ${googleToken}`,
        'Content-Type': 'application/json',
      },
      body: body ? JSON.stringify(body) : undefined,
    },
  );
  const text = await res.text();
  let json;
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    json = { raw: text };
  }
  if (res.status === 401) {
    googleToken = await getGoogleToken();
    return sheets(method, range, body);
  }
  if (!res.ok) {
    throw new Error(`Sheets ${method} ${range} failed: ${JSON.stringify(json)}`);
  }
  return json;
}

async function readRows() {
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}/values/${sheetRange('A:M')}`,
    { headers: { Authorization: `Bearer ${googleToken}` } },
  );
  const json = await res.json();
  if (!res.ok) throw new Error(`Read Sheet failed: ${JSON.stringify(json)}`);
  return json.values || [];
}

async function updateRow(rowNumber, values) {
  return sheets('PUT', `B${rowNumber}:M${rowNumber}`, { values: [values] });
}

async function updateMeta(rowNumber, status, jobId, message, submittedAt, completedAt, attempts, telegramSent) {
  return updateRow(rowNumber, [
    '', '', '', '',
    status || '',
    jobId || '',
    '',
    message || '',
    submittedAt || '',
    completedAt || '',
    attempts ?? '',
    telegramSent ?? '',
  ]);
}

function asImageCell(url) {
  if (!url) return '';
  const escaped = String(url).replace(/"/g, '""');
  return WRITE_IMAGE_FORMULAS ? `=IMAGE("${escaped}")` : url;
}

function normalizePrompt(prompt) {
  let output = String(prompt || '').trim();
  const warnings = [];
  const errors = [];

  if (!output) errors.push('Prompt is empty');

  if (AUTO_FIX_SINGLE_DASH_OW) {
    const fixed = output.replace(/(^|\s)-ow(\s+\d+)/g, '$1--ow$2');
    if (fixed !== output) {
      output = fixed;
      warnings.push('Fixed single-dash -ow to --ow');
    }
  }

  const iwRegex = /--iw\s+([0-9.]+)/g;
  for (const match of output.matchAll(iwRegex)) {
    const value = Number(match[1]);
    if (Number.isFinite(value) && value > 3) {
      errors.push(`Invalid --iw ${match[1]}: Midjourney image weight must be 0-3. Use --ow for Omni Reference weight.`);
    }
  }

  if (/https:\/\/ibb\.co\//i.test(output)) {
    warnings.push('Prompt contains ibb.co page URL; direct i.ibb.co image URLs are safer for API generation.');
  }

  return { prompt: output, warnings, errors };
}

async function submitTTAPI(prompt) {
  const res = await fetch('https://api.ttapi.io/midjourney/v1/imagine', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'TT-API-KEY': TTAPI_KEY,
      'User-Agent': 'n8n-ttapi-v1/1.0',
    },
    body: JSON.stringify({ prompt, mode: MODE, timeout: TIMEOUT_SECONDS }),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok || json.status === 'FAILED' || !json.data?.jobId) {
    throw new Error(json.message || JSON.stringify(json) || `TTAPI submit HTTP ${res.status}`);
  }
  return json.data.jobId;
}

async function fetchTTAPI(jobId) {
  const res = await fetch(`https://api.ttapi.io/midjourney/v1/fetch?jobId=${encodeURIComponent(jobId)}`, {
    headers: {
      'TT-API-KEY': TTAPI_KEY,
      'User-Agent': 'n8n-ttapi-v1/1.0',
    },
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(json.message || `TTAPI fetch HTTP ${res.status}`);
  return json;
}

function safeName(value) {
  return String(value || '').replace(/[^a-zA-Z0-9_.-]+/g, '_').slice(0, 100);
}

async function downloadImage(url, rowNumber, jobId, index) {
  if (!url) return '';
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const ext = path.extname(new URL(url).pathname) || '.png';
  const file = path.join(OUTPUT_DIR, `row_${rowNumber}_job_${safeName(jobId)}_${index}${ext}`);
  const res = await fetch(url);
  if (!res.ok) throw new Error(`download failed ${res.status}: ${url}`);
  const buffer = Buffer.from(await res.arrayBuffer());
  fs.writeFileSync(file, buffer);
  return file;
}

async function sendTelegramDocument(url, caption) {
  if (!SEND_TELEGRAM || !TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID || !url) {
    return { skipped: true };
  }
  const res = await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID,
      document: url,
      caption: caption.slice(0, 1000),
    }),
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok || !json.ok) {
    throw new Error(`Telegram send failed: ${JSON.stringify(json)}`);
  }
  return json;
}

function rowStatus(row) {
  return String(row[5] || '').trim().toUpperCase();
}

function isImageDoneCell(value) {
  const text = String(value || '').trim();
  return text && !/^ERROR$/i.test(text) && !/^FAILED/i.test(text);
}

async function ensureHeaders() {
  await sheets('PUT', 'F1:M1', {
    values: [[
      'STATUS',
      'TTAPI_JOB_ID',
      'GRID_IMAGE_URL',
      'ERROR_OR_PROGRESS',
      'SUBMITTED_AT',
      'COMPLETED_AT',
      'ATTEMPTS',
      'TELEGRAM_SENT',
    ]],
  });
}

function buildQueue(rows) {
  const running = [];
  const ready = [];
  for (let i = 1; i < rows.length; i++) {
    const rowNumber = i + 1;
    const row = rows[i] || [];
    const prompt = String(row[0] || '').trim();
    if (!prompt) continue;
    const image1 = row[1];
    const status = rowStatus(row);
    const jobId = String(row[6] || '').trim();
    const attempts = Number(row[11] || 0) || 0;

    if ((status === 'SUBMITTED' || status === 'RUNNING') && jobId) {
      running.push({ rowNumber, prompt, jobId, attempts, submittedAt: row[9] || '' });
      continue;
    }

    if (status === 'DONE' || isImageDoneCell(image1)) continue;
    if (status === 'LOCKED') continue;

    if (!status || ['READY', 'FAILED', 'FAILED_PROMPT', 'ERROR', 'RETRY'].includes(status) || /^ERROR$/i.test(String(image1 || ''))) {
      ready.push({ rowNumber, prompt, attempts });
    }
  }

  if (MAX_ROWS_PER_RUN.toUpperCase() !== 'ALL') {
    const limit = Math.max(0, Number(MAX_ROWS_PER_RUN) || 0);
    return { running, ready: limit > 0 ? ready.slice(0, limit) : [] };
  }

  return { running, ready };
}

await ensureHeaders();

const startedAt = Date.now();
const deadline = startedAt + MAX_RUNTIME_MINUTES * 60 * 1000;
let rows = await readRows();
let { running, ready } = buildQueue(rows);
const summary = {
  workflow: 'ttapi_v1',
  mode: MODE,
  maxInFlight: MAX_IN_FLIGHT,
  submitted: 0,
  resumed: running.length,
  completed: 0,
  failed: 0,
  promptFailed: 0,
  telegramSent: 0,
  downloaded: 0,
  leftReady: ready.length,
  outputDir: OUTPUT_DIR,
  events: [],
};

async function markFailed(job, status, error) {
  summary.failed += status === 'FAILED_PROMPT' ? 0 : 1;
  summary.promptFailed += status === 'FAILED_PROMPT' ? 1 : 0;
  await updateRow(job.rowNumber, [
    '', '', '', '',
    status,
    job.jobId || '',
    '',
    String(error || '').slice(0, 45000),
    job.submittedAt || '',
    isoNow(),
    job.attempts ?? '',
    '',
  ]);
}

while (Date.now() < deadline && (ready.length > 0 || running.length > 0)) {
  while (running.length < MAX_IN_FLIGHT && ready.length > 0 && Date.now() < deadline) {
    const job = ready.shift();
    const submittedAt = isoNow();
    const attempts = (job.attempts || 0) + 1;
    await updateMeta(job.rowNumber, 'LOCKED', '', 'Preparing prompt', submittedAt, '', attempts, '');

    const normalized = normalizePrompt(job.prompt);
    if (normalized.errors.length > 0) {
      await markFailed({ ...job, attempts, submittedAt }, 'FAILED_PROMPT', normalized.errors.join('; '));
      summary.events.push({ row: job.rowNumber, status: 'FAILED_PROMPT', error: normalized.errors.join('; ') });
      continue;
    }

    try {
      const jobId = await submitTTAPI(normalized.prompt);
      const message = normalized.warnings.length ? normalized.warnings.join('; ') : 'Submitted';
      await updateRow(job.rowNumber, [
        '', '', '', '',
        'SUBMITTED',
        jobId,
        '',
        message,
        submittedAt,
        '',
        attempts,
        '',
      ]);
      running.push({ ...job, prompt: normalized.prompt, jobId, attempts, submittedAt });
      summary.submitted += 1;
      summary.events.push({ row: job.rowNumber, status: 'SUBMITTED', jobId });
    } catch (error) {
      await markFailed({ ...job, attempts, submittedAt }, 'FAILED', error.message);
      summary.events.push({ row: job.rowNumber, status: 'FAILED', error: error.message });
    }
  }

  if (running.length === 0) break;

  await sleep(POLL_SECONDS * 1000);

  const stillRunning = [];
  for (const job of running) {
    try {
      const result = await fetchTTAPI(job.jobId);
      const status = result.status;
      const data = result.data || {};

      if (status === 'SUCCESS') {
        const imageUrls = (data.images || []).filter(Boolean);
        const gridUrl = data.cdnImage || data.discordImage || '';
        const localFiles = [];
        let telegramOk = 0;

        for (let i = 0; i < imageUrls.length; i++) {
          try {
            const file = await downloadImage(imageUrls[i], job.rowNumber, job.jobId, i + 1);
            localFiles.push(file);
            summary.downloaded += 1;
          } catch (error) {
            summary.events.push({ row: job.rowNumber, status: 'DOWNLOAD_ERROR', error: error.message });
          }

          try {
            await sendTelegramDocument(imageUrls[i], `Row ${job.rowNumber} image ${i + 1}/${imageUrls.length}`);
            telegramOk += 1;
          } catch (error) {
            summary.events.push({ row: job.rowNumber, status: 'TELEGRAM_ERROR', error: error.message });
          }
        }

        await updateRow(job.rowNumber, [
          asImageCell(imageUrls[0]),
          asImageCell(imageUrls[1]),
          asImageCell(imageUrls[2]),
          asImageCell(imageUrls[3]),
          telegramOk === imageUrls.length ? 'DONE' : 'DONE_TELEGRAM_ERROR',
          job.jobId,
          gridUrl,
          localFiles.join('\n'),
          job.submittedAt,
          isoNow(),
          job.attempts,
          `${telegramOk}/${imageUrls.length}`,
        ]);
        summary.completed += 1;
        summary.telegramSent += telegramOk;
        summary.events.push({ row: job.rowNumber, status: 'DONE', jobId: job.jobId, images: imageUrls.length });
      } else if (status === 'FAILED' || status === 'FAILURE') {
        const message = result.message || data.error || data.failReason || 'TTAPI failed';
        await markFailed(job, 'FAILED', message);
        summary.events.push({ row: job.rowNumber, status: 'FAILED', error: message });
      } else {
        const progress = data.progress ? `Progress ${data.progress}%` : status || 'RUNNING';
        await updateRow(job.rowNumber, [
          '', '', '', '',
          'RUNNING',
          job.jobId,
          '',
          progress,
          job.submittedAt,
          '',
          job.attempts,
          '',
        ]);
        stillRunning.push(job);
      }
    } catch (error) {
      await updateRow(job.rowNumber, [
        '', '', '', '',
        'RUNNING',
        job.jobId,
        '',
        `Poll error, will retry: ${error.message}`.slice(0, 45000),
        job.submittedAt,
        '',
        job.attempts,
        '',
      ]);
      stillRunning.push(job);
    }
  }
  running = stillRunning;
}

summary.leftRunning = running.length;
summary.leftReady = ready.length;
summary.finishedAt = isoNow();

return [{ json: summary }];
"""


def get_source_config(conn):
    row = conn.execute(
        "SELECT nodes FROM workflow_entity WHERE id = ?", (SOURCE_WORKFLOW_ID,)
    ).fetchone()
    if not row:
        raise SystemExit(f"Source workflow not found: {SOURCE_WORKFLOW_ID}")

    nodes = json.loads(row[0])
    legacy_credentials = ""
    ttapi_key = ""
    for node in nodes:
        if node.get("name") == "Edit Fields":
            assignments = (
                node.get("parameters", {})
                .get("assignments", {})
                .get("assignments", [])
            )
            for assignment in assignments:
                if assignment.get("name") == "✅CREDENTIALS":
                    legacy_credentials = assignment.get("value", "")
        if node.get("name") == "Generate Image":
            params = (
                node.get("parameters", {})
                .get("headerParameters", {})
                .get("parameters", [])
            )
            for header in params:
                if header.get("name") == "TT-API-KEY":
                    ttapi_key = str(header.get("value", "")).lstrip("=")

    if not legacy_credentials:
        raise SystemExit("Could not extract legacy credentials from source workflow")
    if not ttapi_key:
        raise SystemExit("Could not extract TTAPI key from source workflow")
    return legacy_credentials, ttapi_key


def build_workflow_nodes(legacy_credentials, ttapi_key):
    manual_id = str(uuid.uuid4())
    webhook_id = str(uuid.uuid4())
    config_id = str(uuid.uuid4())
    worker_id = str(uuid.uuid4())
    note_id = str(uuid.uuid4())

    return [
        {
            "parameters": {
                "content": (
                    "## ttapi_v1\n"
                    "Manual TTAPI Relax batch runner.\n\n"
                    "Sheet layout: A=prompt, B:E=image outputs, F:M=status metadata.\n"
                    "Uses TTAPI relax mode, max 5 concurrent jobs, 1200s timeout, "
                    "downloads files locally, and sends image documents to Telegram.\n\n"
                    "Run with the Manual Trigger. Original workflow is untouched."
                ),
                "height": 260,
                "width": 420,
                "color": 4,
            },
            "type": "n8n-nodes-base.stickyNote",
            "typeVersion": 1,
            "position": [0, -260],
            "id": note_id,
            "name": "README",
        },
        {
            "parameters": {},
            "type": "n8n-nodes-base.manualTrigger",
            "typeVersion": 1,
            "position": [0, 120],
            "id": manual_id,
            "name": "Manual Start",
        },
        {
            "parameters": {
                "httpMethod": "POST",
                "path": "ttapi-v1-run",
                "options": {},
            },
            "type": "n8n-nodes-base.webhook",
            "typeVersion": 1.1,
            "position": [0, 320],
            "id": webhook_id,
            "name": "Webhook Start",
            "webhookId": str(uuid.uuid4()),
        },
        {
            "parameters": {
                "assignments": {
                    "assignments": [
                        {
                            "id": str(uuid.uuid4()),
                            "name": "LEGACY_CREDENTIALS",
                            "value": legacy_credentials,
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "TTAPI_API_KEY",
                            "value": ttapi_key,
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "SHEET_NAME",
                            "value": SHEET_NAME,
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "MODE",
                            "value": "relax",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "TIMEOUT_SECONDS",
                            "value": "1200",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "MAX_IN_FLIGHT",
                            "value": "5",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "POLL_SECONDS",
                            "value": "30",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "MAX_RUNTIME_MINUTES",
                            "value": "55",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "MAX_ROWS_PER_RUN",
                            "value": "ALL",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "SEND_TELEGRAM",
                            "value": "true",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "WRITE_IMAGE_FORMULAS",
                            "value": "true",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "AUTO_FIX_SINGLE_DASH_OW",
                            "value": "true",
                            "type": "string",
                        },
                        {
                            "id": str(uuid.uuid4()),
                            "name": "OUTPUT_DIR",
                            "value": OUTPUT_DIR,
                            "type": "string",
                        },
                    ]
                },
                "options": {},
            },
            "type": "n8n-nodes-base.set",
            "typeVersion": 3.4,
            "position": [280, 220],
            "id": config_id,
            "name": "Edit Config",
        },
        {
            "parameters": {
                "mode": "runOnceForAllItems",
                "jsCode": WORKER_JS,
            },
            "type": "n8n-nodes-base.code",
            "typeVersion": 2,
            "position": [560, 220],
            "id": worker_id,
            "name": "TTAPI Batch Worker",
        },
    ], {
        "Manual Start": {
            "main": [[{"node": "Edit Config", "type": "main", "index": 0}]]
        },
        "Webhook Start": {
            "main": [[{"node": "Edit Config", "type": "main", "index": 0}]]
        },
        "Edit Config": {
            "main": [[{"node": "TTAPI Batch Worker", "type": "main", "index": 0}]]
        },
    }


def main():
    conn = sqlite3.connect(DB_PATH)
    try:
        legacy_credentials, ttapi_key = get_source_config(conn)
        existing = conn.execute(
            "SELECT id FROM workflow_entity WHERE name = ?", (NEW_WORKFLOW_NAME,)
        ).fetchone()
        if existing:
            raise SystemExit(
                f"Workflow named {NEW_WORKFLOW_NAME!r} already exists: {existing[0]}"
            )

        source_share = conn.execute(
            "SELECT projectId, role FROM shared_workflow WHERE workflowId = ?",
            (SOURCE_WORKFLOW_ID,),
        ).fetchone()
        if not source_share:
            raise SystemExit("Source workflow share/project row not found")

        workflow_id = n8n_id()
        nodes, connections = build_workflow_nodes(legacy_credentials, ttapi_key)
        created = now_sql()

        conn.execute("BEGIN")
        conn.execute(
            """
            INSERT INTO workflow_entity (
                id, name, active, nodes, connections, settings, staticData,
                pinData, versionId, triggerCount, meta, parentFolderId,
                createdAt, updatedAt, isArchived, versionCounter, description,
                activeVersionId
            )
            VALUES (?, ?, 0, ?, ?, ?, NULL, NULL, ?, 0, ?, NULL, ?, ?, 0, 1, ?, NULL)
            """,
            (
                workflow_id,
                NEW_WORKFLOW_NAME,
                json.dumps(nodes, separators=(",", ":")),
                json.dumps(connections, separators=(",", ":")),
                json.dumps({"executionOrder": "v1"}, separators=(",", ":")),
                str(uuid.uuid4()),
                json.dumps({"templateCredsSetupCompleted": True}, separators=(",", ":")),
                created,
                created,
                "TTAPI Relax batch runner created from MIDJOURNEY_midapi.ai_v2. Original workflow untouched.",
            ),
        )
        conn.execute(
            """
            INSERT INTO shared_workflow (workflowId, projectId, role, createdAt, updatedAt)
            VALUES (?, ?, ?, ?, ?)
            """,
            (workflow_id, source_share[0], source_share[1], created, created),
        )
        conn.commit()
        print(json.dumps({"workflowId": workflow_id, "name": NEW_WORKFLOW_NAME}))
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
