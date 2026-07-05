// REIMBURSEMENT_27_EXPORT v2
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
const maxDaysPerReport = Math.min(Math.max(Number(cfg['✅REIMBURSEMENT_MAX_DAYS_PER_REPORT'] || 7) || 7, 1), 30);
const pollDelayMs = Math.min(Math.max(Number(cfg['✅REIMBURSEMENT_POLL_DELAY_MS'] || 15000) || 15000, 5000), 180000);
const maxPolls = Math.min(Math.max(Number(cfg['✅REIMBURSEMENT_MAX_POLLS'] || 20) || 20, 1), 100);

const defaultColumns = [
  'RunDate',
  'StartDate',
  'EndDate',
  'ReportStartDate',
  'ReportEndDate',
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
  const cols = raw.split(/\r?\n/).map((v) => v.trim()).filter(Boolean);
  return cols.length ? cols : defaultColumns;
}

const columns = configuredColumns();

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseDate(value) {
  const [y, m, d] = value.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

function formatDate(value) {
  return value.toISOString().slice(0, 10);
}

function addDays(value, days) {
  const copy = new Date(value.getTime());
  copy.setUTCDate(copy.getUTCDate() + days);
  return copy;
}

function buildChunks(start, end) {
  const startDt = parseDate(start);
  const endDt = parseDate(end);
  if (startDt > endDt) throw new Error('START_DATE must be before or equal to END_DATE');
  const chunks = [];
  for (let cursor = startDt; cursor <= endDt;) {
    let chunkEnd = addDays(cursor, maxDaysPerReport - 1);
    if (chunkEnd > endDt) chunkEnd = endDt;
    chunks.push({ start: formatDate(cursor), end: formatDate(chunkEnd) });
    cursor = addDays(chunkEnd, 1);
  }
  return chunks;
}

function rangeFor(range) {
  return encodeURIComponent(tabName + '!' + range);
}

function qSheet(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
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

async function ensureSheet() {
  const meta = await googleRequest.call(this, 'GET', '?fields=sheets.properties.title');
  const existingTabs = new Set((meta.sheets || []).map((s) => s.properties.title));
  let sheetCreated = false;
  if (!existingTabs.has(tabName)) {
    await googleRequest.call(this, 'POST', ':batchUpdate', {
      requests: [{ addSheet: { properties: { title: tabName } } }],
    });
    sheetCreated = true;
  }
  return sheetCreated;
}

async function replaceSheet(values) {
  await googleRequest.call(this, 'POST', '/values:batchClear', {
    ranges: [qSheet(tabName) + '!A:ZZ'],
  });
  const chunkSize = 1000;
  for (let offset = 0; offset < values.length; offset += chunkSize) {
    const chunk = values.slice(offset, offset + chunkSize);
    const rowNumber = 1 + offset;
    await googleRequest.call(this, 'PUT', '/values/' + rangeFor('A' + rowNumber) + '?valueInputOption=RAW', {
      values: chunk,
    });
  }
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
        Accept: 'application/json',
      },
    };
    if (body !== undefined) options.body = body;
    return await this.helpers.httpRequest(options);
  } catch (error) {
    throw normalizeHttpError(error, method + ' ' + path);
  }
}

async function createReport(chunk) {
  const response = await spApi.call(this, 'POST', '/reports/2021-06-30/reports', {
    reportType,
    dataStartTime: chunk.start + 'T00:00:00Z',
    dataEndTime: chunk.end + 'T23:59:59Z',
    marketplaceIds: [marketplaceId],
  });
  const reportId = response.reportId || response.payload?.reportId;
  if (!reportId) throw new Error('Amazon createReport did not return reportId: ' + JSON.stringify(response).slice(0, 500));
  return reportId;
}

async function getReport(reportId) {
  return await spApi.call(this, 'GET', '/reports/2021-06-30/reports/' + encodeURIComponent(reportId));
}

async function findExistingDoneReport(chunk) {
  const query = [
    'reportTypes=' + encodeURIComponent(reportType),
    'processingStatuses=DONE',
    'pageSize=100',
  ].join('&');
  const response = await spApi.call(this, 'GET', '/reports/2021-06-30/reports?' + query);
  const reports = response.reports || [];
  const startPrefix = chunk.start + 'T00:00:00';
  const endPrefix = chunk.end + 'T23:59:59';
  return reports.find((report) => {
    const dataStart = String(report.dataStartTime || '');
    const dataEnd = String(report.dataEndTime || '');
    return dataStart.startsWith(startPrefix) && dataEnd.startsWith(endPrefix) && report.reportDocumentId;
  }) || null;
}

async function getReportDocument(reportDocumentId) {
  return await spApi.call(this, 'GET', '/reports/2021-06-30/documents/' + encodeURIComponent(reportDocumentId));
}

async function downloadReportText(documentInfo) {
  const url = documentInfo.url;
  if (!url) throw new Error('Report document missing download URL');
  let response;
  try {
    response = await this.helpers.httpRequest({ method: 'GET', url, encoding: 'arraybuffer', returnFullResponse: true });
  } catch (error) {
    throw normalizeHttpError(error, 'Download report document');
  }
  let buffer = Buffer.from(response.body || '');
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

function rowFromObject(obj) {
  return columns.map((column) => obj[column] ?? '');
}

function buildOutputRows(parsed, reportId, reportDocumentId, chunk) {
  return parsed.rows.map((row, index) => {
    const dedupKey = [
      row['reimbursement-id'] || '',
      row['case-id'] || '',
      row['amazon-order-id'] || '',
      row.sku || '',
      row.fnsku || '',
      row.reason || '',
    ].join('|');
    return rowFromObject({
      RunDate: runDate,
      StartDate: startDate,
      EndDate: endDate,
      ReportStartDate: chunk.start,
      ReportEndDate: chunk.end,
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
    });
  });
}

function statusRow(chunk, reportId, status, note) {
  return rowFromObject({
    RunDate: runDate,
    StartDate: startDate,
    EndDate: endDate,
    ReportStartDate: chunk.start,
    ReportEndDate: chunk.end,
    MarketplaceId: marketplaceId,
    ReportType: reportType,
    ReportId: reportId,
    ProcessingStatus: status,
    DownloadedAt: new Date().toISOString(),
    PossibleFollowUp: note,
  });
}

function terminal(status) {
  return ['DONE', 'CANCELLED', 'FATAL'].includes(String(status || '').toUpperCase());
}

function noteForNonDoneStatus(status) {
  if (status === 'CANCELLED') {
    return 'Amazon cancelled this chunk with no downloadable document. For this report type, that often means no report rows were available for the period; re-run later if you need to verify.';
  }
  if (status === 'FATAL') {
    return 'Amazon returned FATAL for this chunk. Reduce ✅REIMBURSEMENT_MAX_DAYS_PER_REPORT or re-run later.';
  }
  return 'Chunk did not produce a downloadable reimbursement report. Re-run later or increase ✅REIMBURSEMENT_MAX_POLLS.';
}

async function runChunk(chunk) {
  let reusedExistingReport = false;
  let existingReport = chunk.existingReport || await findExistingDoneReport.call(this, chunk);
  let reportId = existingReport?.reportId;
  let report = null;
  if (existingReport) {
    report = existingReport;
    reusedExistingReport = true;
  } else {
    reportId = await createReport.call(this, chunk);
    for (let attempt = 1; attempt <= maxPolls; attempt++) {
      report = await getReport.call(this, reportId);
      const status = String(report.processingStatus || '').toUpperCase();
      if (terminal(status)) break;
      await sleep(pollDelayMs);
    }
  }

  const finalStatus = String(report?.processingStatus || 'UNKNOWN').toUpperCase();
  if (finalStatus !== 'DONE') {
    existingReport = await findExistingDoneReport.call(this, chunk);
    if (existingReport) {
      return await runChunk.call(this, { ...chunk, existingReport });
    }
    return {
      status: finalStatus,
      reportId,
      reportDocumentId: '',
      rows: [],
      statusRow: statusRow(chunk, reportId, finalStatus, noteForNonDoneStatus(finalStatus)),
    };
  }

  const reportDocumentId = report.reportDocumentId;
  if (!reportDocumentId) throw new Error('Report DONE but missing reportDocumentId: ' + JSON.stringify(report).slice(0, 500));
  const documentInfo = await getReportDocument.call(this, reportDocumentId);
  const reportText = await downloadReportText.call(this, documentInfo);
  const parsed = parseTsv(reportText);
  return {
    status: 'DONE',
    reportId,
    reportDocumentId,
    reusedExistingReport,
    sourceColumns: parsed.headers,
    rows: buildOutputRows(parsed, reportId, reportDocumentId, chunk),
    rowCount: parsed.rows.length,
  };
}

const sheetCreated = await ensureSheet.call(this);
const exactExistingReport = await findExistingDoneReport.call(this, { start: startDate, end: endDate });
const chunks = exactExistingReport
  ? [{ start: startDate, end: endDate, existingReport: exactExistingReport }]
  : buildChunks(startDate, endDate);
const outputRows = [];
const chunkSummaries = [];
const sourceColumns = new Set();

for (const chunk of chunks) {
  const result = await runChunk.call(this, chunk);
  chunkSummaries.push({
    start: chunk.start,
    end: chunk.end,
    status: result.status,
    reportId: result.reportId,
    rows: result.rowCount || 0,
    reusedExistingReport: !!result.reusedExistingReport,
  });
  for (const col of result.sourceColumns || []) sourceColumns.add(col);
  if (result.rows.length) outputRows.push(...result.rows);
  if (result.status !== 'DONE') outputRows.push(result.statusRow);
}

if (!outputRows.length) {
  outputRows.push(statusRow(
    { start: startDate, end: endDate },
    '',
    'DONE_NO_ROWS',
    'Amazon returned no reimbursement rows for this date range.'
  ));
}

const fatalCount = chunkSummaries.filter((c) => c.status === 'FATAL').length;
const cancelledCount = chunkSummaries.filter((c) => c.status === 'CANCELLED').length;
const nonDoneCount = chunkSummaries.filter((c) => c.status !== 'DONE').length;
const dataRows = outputRows.length - nonDoneCount;
let finalStatus = 'done';
if (fatalCount > 0) finalStatus = dataRows > 0 ? 'partial_with_fatal_chunks' : 'no_downloadable_chunks';
else if (cancelledCount > 0) finalStatus = dataRows > 0 ? 'done_with_cancelled_empty_chunks' : 'done_no_rows_or_cancelled';

const preservedExistingSheet = dataRows === 0 && fatalCount === chunks.length && chunks.length > 0;
if (!preservedExistingSheet) {
  await replaceSheet.call(this, [columns, ...outputRows]);
} else {
  finalStatus = 'preserved_existing_sheet_due_to_all_fatal_chunks';
}

return [{
  json: {
    status: finalStatus,
    tabName,
    sheetCreated,
    reportType,
    chunksRequested: chunks.length,
    maxDaysPerReport,
    pollDelayMs,
    maxPolls,
    fatalChunks: fatalCount,
    cancelledChunks: cancelledCount,
    rowsWritten: outputRows.length,
    dataRows,
    preservedExistingSheet,
    sourceColumns: Array.from(sourceColumns),
    dateRange: startDate + ' to ' + endDate,
    chunkSummaries,
  },
}];
