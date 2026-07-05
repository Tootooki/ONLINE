import json
from pathlib import Path


SOURCE = Path("/tmp/n8n_workflow_vEN15jkOLWmpAXxn.json")
OUTPUT = Path("/tmp/n8n_workflow_vEN15jkOLWmpAXxn_patched.json")


REIMBURSEMENT_CODE = r"""
// REIMBURSEMENT_27_EXPORT v1
// Uses the dynamic credential block parsed by SETUP/MERGE.
// Amazon side: requests/downloads report data only. No cases, feeds, or Seller Central mutations.
// Google side: creates/updates the configured reimbursement tab with the latest report output.

const zlib = require('zlib');

const cfg = $input.first().json;
const runFlag = String(cfg['✅GOOGLE_SHEET_27_REIMBURSEMENT_RUN'] || 'YES').trim().toUpperCase();
if (!['YES', 'Y', 'TRUE', '1', 'ON'].includes(runFlag)) {
  return [{ json: { status: 'skipped', reason: '✅GOOGLE_SHEET_27_REIMBURSEMENT_RUN is not YES' } }];
}

const spreadsheetId = String(cfg['✅GOOGLE_SHEET_ID'] || '').trim();
const googleToken = String(cfg._google_token || '').trim();
const tabName = String(cfg['✅GOOGLE_SHEET_27_REIMBURSEMENT'] || '27_REIMBURSEMENT').trim();
const marketplaceId = String(cfg['✅AMZ_SP_MARKETPLACE_ID'] || '').trim();
const spAccessToken = String(cfg.sp_access_token || '').trim();
const startDate = String(cfg['✅START_DATE'] || '').trim();
const endDate = String(cfg['✅END_DATE'] || '').trim();

if (!spreadsheetId) throw new Error('Missing ✅GOOGLE_SHEET_ID');
if (!googleToken) throw new Error('Missing _google_token from SETUP');
if (!tabName) throw new Error('Missing ✅GOOGLE_SHEET_27_REIMBURSEMENT');
if (!marketplaceId) throw new Error('Missing ✅AMZ_SP_MARKETPLACE_ID');
if (!spAccessToken) throw new Error('Missing sp_access_token from MERGE');
if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate) || !/^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
  throw new Error('START_DATE and END_DATE must be YYYY-MM-DD');
}

const reportType = 'GET_FBA_REIMBURSEMENTS_DATA';
const spApiBase = String(cfg['✅AMZ_SP_API_BASE_URL'] || 'https://sellingpartnerapi-na.amazon.com').replace(/\/+$/, '');
const runDate = new Date().toISOString();

const defaultColumns = [
  'RunDate',
  'StartDate',
  'EndDate',
  'MarketplaceId',
  'ReportType',
  'ReportId',
  'ProcessingStatus',
  'ReportDocumentId',
  'DownloadedAt',
  'SourceRowNumber',
  'DedupKey',
  'SuggestedTopic',
  'PossibleFollowUp',
  'approval-date',
  'reimbursement-id',
  'case-id',
  'amazon-order-id',
  'reason',
  'sku',
  'fnsku',
  'asin',
  'product-name',
  'condition',
  'currency-unit',
  'amount-per-unit',
  'amount-total',
  'quantity-reimbursed-cash',
  'quantity-reimbursed-inventory',
  'quantity-reimbursed-total',
  'original-reimbursement-id',
  'original-reimbursement-type',
  'raw_json'
];

function configuredColumns() {
  const raw = String(cfg['✅GOOGLE_SHEET_27_REIMBURSEMENT_COLUMNS'] || '').trim();
  if (!raw) return defaultColumns;
  const cols = raw
    .split(/\r?\n/)
    .map((v) => v.trim())
    .filter(Boolean);
  return cols.length ? cols : defaultColumns;
}

const columns = configuredColumns();

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function dateTimeStart(d) {
  return d + 'T00:00:00Z';
}

function dateTimeEnd(d) {
  return d + 'T23:59:59Z';
}

async function googleRequest(method, path, body) {
  const options = {
    method,
    url: 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + path,
    headers: { Authorization: 'Bearer ' + googleToken, 'Content-Type': 'application/json' },
  };
  if (body !== undefined) options.body = JSON.stringify(body);
  return await this.helpers.httpRequest(options);
}

async function ensureSheetAndHeaders() {
  const meta = await googleRequest.call(this, 'GET', '?fields=sheets.properties.title');
  const existingTabs = new Set((meta.sheets || []).map((s) => s.properties.title));
  let sheetCreated = false;
  if (!existingTabs.has(tabName)) {
    await googleRequest.call(this, 'POST', ':batchUpdate', {
      requests: [{ addSheet: { properties: { title: tabName } } }],
    });
    sheetCreated = true;
  }
  const encodedTab = encodeURIComponent(tabName);
  await googleRequest.call(this, 'POST', '/values/' + encodedTab + '!A:ZZ:clear', {});
  await googleRequest.call(this, 'PUT', '/values/' + encodedTab + '!A1?valueInputOption=RAW', { values: [columns] });
  return sheetCreated;
}

function normalizeHttpError(error, context) {
  const body = error?.response?.body || error?.cause?.response?.body || error?.message || String(error);
  return new Error(context + ': ' + (typeof body === 'string' ? body.slice(0, 1000) : JSON.stringify(body).slice(0, 1000)));
}

async function spApi(method, path, body) {
  try {
    const options = {
      method,
      url: spApiBase + path,
      headers: {
        'x-amz-access-token': spAccessToken,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
    };
    if (body !== undefined) options.body = JSON.stringify(body);
    return await this.helpers.httpRequest(options);
  } catch (error) {
    throw normalizeHttpError(error, method + ' ' + path);
  }
}

async function createReport() {
  const response = await spApi.call(this, 'POST', '/reports/2021-06-30/reports', {
    reportType,
    dataStartTime: dateTimeStart(startDate),
    dataEndTime: dateTimeEnd(endDate),
    marketplaceIds: [marketplaceId],
  });
  const reportId = response.reportId || response.payload?.reportId;
  if (!reportId) throw new Error('Amazon createReport did not return reportId: ' + JSON.stringify(response).slice(0, 500));
  return reportId;
}

async function getReport(reportId) {
  return await spApi.call(this, 'GET', '/reports/2021-06-30/reports/' + encodeURIComponent(reportId));
}

async function getReportDocument(reportDocumentId) {
  return await spApi.call(this, 'GET', '/reports/2021-06-30/documents/' + encodeURIComponent(reportDocumentId));
}

async function downloadReportText(documentInfo) {
  const url = documentInfo.url;
  if (!url) throw new Error('Report document missing download URL');
  let body;
  try {
    body = await this.helpers.httpRequest({ method: 'GET', url, encoding: 'arraybuffer', json: false });
  } catch (error) {
    throw normalizeHttpError(error, 'Download report document');
  }
  let buffer;
  if (Buffer.isBuffer(body)) buffer = body;
  else if (body instanceof ArrayBuffer) buffer = Buffer.from(body);
  else if (ArrayBuffer.isView(body)) buffer = Buffer.from(body.buffer);
  else if (typeof body === 'string') buffer = Buffer.from(body, 'binary');
  else buffer = Buffer.from(String(body || ''), 'binary');

  if (String(documentInfo.compressionAlgorithm || '').toUpperCase() === 'GZIP') {
    buffer = zlib.gunzipSync(buffer);
  }
  return buffer.toString('utf8').replace(/^\uFEFF/, '');
}

function parseTsv(text) {
  const clean = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = clean.split('\n').filter((line) => line.length);
  if (!lines.length) return { headers: [], rows: [] };
  const headers = lines[0].split('\t').map((h) => h.trim());
  const rows = lines.slice(1).map((line) => {
    const values = line.split('\t');
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = values[index] ?? '';
    });
    return obj;
  });
  return { headers, rows };
}

function suggestedTopic(row) {
  const reason = String(row.reason || row['reason-code'] || '').toLowerCase();
  if (reason.includes('lost') || reason.includes('missing')) return 'FC_LOST_OR_MISSING_CONFIRMED';
  if (reason.includes('damage')) return 'FC_DAMAGED_CONFIRMED';
  if (reason.includes('customer') || reason.includes('return')) return 'CUSTOMER_RETURN_REIMBURSED';
  if (reason.includes('removal')) return 'REMOVAL_REIMBURSED';
  if (reason.includes('warehouse')) return 'FC_OPERATION_REIMBURSED';
  return 'REIMBURSEMENT_HISTORY';
}

function possibleFollowUp(row) {
  const amount = Number(row['amount-total'] || row['amount-per-unit'] || 0);
  const originalId = String(row['original-reimbursement-id'] || '').trim();
  if (originalId) return 'Check whether this is a reversal or adjustment against original reimbursement.';
  if (amount === 0) return 'Zero amount row; review cash vs inventory reimbursement fields.';
  return 'Use for dedupe and payment confirmation before creating any new case.';
}

function buildOutputRows(parsed, reportId, reportDocumentId) {
  return parsed.rows.map((row, index) => {
    const dedupKey = [
      row['reimbursement-id'] || '',
      row['case-id'] || '',
      row['amazon-order-id'] || '',
      row.sku || '',
      row.fnsku || '',
      row.reason || '',
    ].join('|');
    const enriched = {
      RunDate: runDate,
      StartDate: startDate,
      EndDate: endDate,
      MarketplaceId: marketplaceId,
      ReportType: reportType,
      ReportId: reportId,
      ProcessingStatus: 'DONE',
      ReportDocumentId: reportDocumentId,
      DownloadedAt: new Date().toISOString(),
      SourceRowNumber: index + 2,
      DedupKey: dedupKey,
      SuggestedTopic: suggestedTopic(row),
      PossibleFollowUp: possibleFollowUp(row),
      ...row,
      raw_json: JSON.stringify(row),
    };
    return columns.map((column) => enriched[column] ?? '');
  });
}

async function writeRows(rows) {
  const encodedTab = encodeURIComponent(tabName);
  if (!rows.length) {
    const statusRow = Object.fromEntries(columns.map((col) => [col, '']));
    statusRow.RunDate = runDate;
    statusRow.StartDate = startDate;
    statusRow.EndDate = endDate;
    statusRow.MarketplaceId = marketplaceId;
    statusRow.ReportType = reportType;
    statusRow.ProcessingStatus = 'DONE_NO_ROWS';
    statusRow.PossibleFollowUp = 'Amazon returned no reimbursement rows for this date range.';
    await googleRequest.call(this, 'PUT', '/values/' + encodedTab + '!A2?valueInputOption=RAW', {
      values: [columns.map((column) => statusRow[column] ?? '')],
    });
    return;
  }

  const chunkSize = 1000;
  for (let offset = 0; offset < rows.length; offset += chunkSize) {
    const chunk = rows.slice(offset, offset + chunkSize);
    const rowNumber = 2 + offset;
    await googleRequest.call(this, 'PUT', '/values/' + encodedTab + '!A' + rowNumber + '?valueInputOption=RAW', {
      values: chunk,
    });
  }
}

function terminal(status) {
  return ['DONE', 'CANCELLED', 'FATAL'].includes(String(status || '').toUpperCase());
}

const sheetCreated = await ensureSheetAndHeaders.call(this);
const reportId = await createReport.call(this);

let report = null;
const pollDelayMs = Number(cfg['✅REIMBURSEMENT_POLL_DELAY_MS'] || 30000);
const maxPolls = Number(cfg['✅REIMBURSEMENT_MAX_POLLS'] || 20);
for (let attempt = 1; attempt <= maxPolls; attempt++) {
  report = await getReport.call(this, reportId);
  const status = String(report.processingStatus || '').toUpperCase();
  if (terminal(status)) break;
  await sleep(pollDelayMs);
}

const finalStatus = String(report?.processingStatus || 'UNKNOWN').toUpperCase();
if (finalStatus !== 'DONE') {
  const encodedTab = encodeURIComponent(tabName);
  const statusRow = Object.fromEntries(columns.map((col) => [col, '']));
  statusRow.RunDate = runDate;
  statusRow.StartDate = startDate;
  statusRow.EndDate = endDate;
  statusRow.MarketplaceId = marketplaceId;
  statusRow.ReportType = reportType;
  statusRow.ReportId = reportId;
  statusRow.ProcessingStatus = finalStatus;
  statusRow.PossibleFollowUp = 'Report was requested but not ready yet. Re-run later or increase ✅REIMBURSEMENT_MAX_POLLS.';
  await googleRequest.call(this, 'PUT', '/values/' + encodedTab + '!A2?valueInputOption=RAW', {
    values: [columns.map((column) => statusRow[column] ?? '')],
  });
  return [{
    json: {
      status: 'requested_not_ready',
      tabName,
      sheetCreated,
      reportType,
      reportId,
      processingStatus: finalStatus,
      maxPolls,
      pollDelayMs,
    },
  }];
}

const reportDocumentId = report.reportDocumentId;
if (!reportDocumentId) throw new Error('Report DONE but missing reportDocumentId: ' + JSON.stringify(report).slice(0, 500));

const documentInfo = await getReportDocument.call(this, reportDocumentId);
const reportText = await downloadReportText.call(this, documentInfo);
const parsed = parseTsv(reportText);
const rows = buildOutputRows(parsed, reportId, reportDocumentId);
await writeRows.call(this, rows);

return [{
  json: {
    status: 'done',
    tabName,
    sheetCreated,
    reportType,
    reportId,
    reportDocumentId,
    sourceColumns: parsed.headers,
    rowsWritten: rows.length,
    dateRange: startDate + ' to ' + endDate,
  },
}];
""".strip()


def main() -> None:
    workflow = json.loads(SOURCE.read_text())

    # Do not duplicate the node if this script is re-run.
    workflow["nodes"] = [
        node for node in workflow["nodes"] if node.get("name") != "REIMBURSEMENT_27_EXPORT"
    ]

    node = {
        "parameters": {"jsCode": REIMBURSEMENT_CODE},
        "id": "7fc89ccb-a6e2-446b-bbb8-c88f0ee9c27b",
        "name": "REIMBURSEMENT_27_EXPORT",
        "type": "n8n-nodes-base.code",
        "typeVersion": 2,
        "position": [1040, 48],
    }
    workflow["nodes"].append(node)

    connections = workflow.setdefault("connections", {})
    connections["MERGE"] = {
        "main": [[{"node": "REIMBURSEMENT_27_EXPORT", "type": "main", "index": 0}]]
    }
    # Remove stale inbound connections to this node if present elsewhere.
    for src, src_conns in list(connections.items()):
        if src == "MERGE":
            continue
        for output_name, outputs in list(src_conns.items()):
            for output in outputs:
                output[:] = [edge for edge in output if edge.get("node") != "REIMBURSEMENT_27_EXPORT"]

    payload = {
        "name": workflow["name"],
        "nodes": workflow["nodes"],
        "connections": workflow["connections"],
        "settings": workflow.get("settings", {}),
    }
    if workflow.get("staticData") is not None:
        payload["staticData"] = workflow.get("staticData")
    if workflow.get("pinData") is not None:
        payload["pinData"] = workflow.get("pinData")

    OUTPUT.write_text(json.dumps(payload, indent=2))
    print(OUTPUT)
    print(f"nodes={len(payload['nodes'])}")


if __name__ == "__main__":
    main()
