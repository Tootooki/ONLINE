import fs from "node:fs/promises";
import crypto from "node:crypto";

const WORKFLOW_JSON = "/tmp/reimb_current_for_actions.json";
const TAB_NAME = "REIMB_ALL_ACTIONS";
const COLUMN_COUNT = 15;

function base64url(input) {
  return Buffer.from(input)
    .toString("base64")
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
}

function parseCredentialTable(text) {
  let src = String(text || "");
  if (src.startsWith("=")) src = src.slice(1);

  const rows = [[]];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < src.length; i += 1) {
    const c = src[i];
    if (inQuotes) {
      if (c === '"') {
        if (src[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
      continue;
    }

    if (c === '"') {
      inQuotes = true;
    } else if (c === "\t") {
      rows[rows.length - 1].push(field);
      field = "";
    } else if (c === "\n") {
      rows[rows.length - 1].push(field);
      rows.push([]);
      field = "";
    } else if (c !== "\r") {
      field += c;
    }
  }
  rows[rows.length - 1].push(field);

  const header = rows.shift() || [];
  const nameIdx = header.findIndex((h) => String(h).trim().toUpperCase() === "NAME");
  const infoIdx = header.findIndex((h) => String(h).trim().toUpperCase() === "INFO");
  if (nameIdx < 0 || infoIdx < 0) {
    throw new Error("Could not parse CREDENTIALS table header.");
  }

  const result = {};
  for (const row of rows) {
    const key = String(row[nameIdx] || "").trim();
    if (!key) continue;
    result[key] = row[infoIdx] ?? "";
  }
  return result;
}

async function loadCredentials() {
  const workflow = JSON.parse(await fs.readFile(WORKFLOW_JSON, "utf8"));
  const node = workflow.nodes.find((n) => n.name === "CREDENTIALS");
  if (!node) throw new Error("CREDENTIALS node not found in workflow JSON.");

  const assignment = node.parameters?.assignments?.assignments?.find((a) => a.name === "✅CREDENTIALS")
    || node.parameters?.assignments?.assignments?.[0];
  if (!assignment?.value) throw new Error("CREDENTIALS assignment is empty.");

  const credentials = parseCredentialTable(assignment.value);
  const findKey = (needle) => {
    const exact = credentials[needle];
    if (exact !== undefined) return exact;
    const normalizedNeedle = needle.replace(/^✅/, "").toLowerCase();
    const found = Object.entries(credentials).find(([key]) =>
      key.replace(/^✅/, "").toLowerCase() === normalizedNeedle
    );
    return found?.[1];
  };

  const spreadsheetId = findKey("✅GOOGLE_SHEET_ID");
  const clientEmail = findKey("✅GOOGLE_SERVICE_ACCOUNT_EMAIL");
  let privateKey = findKey("✅GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY");
  if (privateKey) privateKey = privateKey.replace(/\\n/g, "\n");

  if (!spreadsheetId || !clientEmail || !privateKey) {
    throw new Error("Missing one or more Google Sheets service account fields.");
  }

  return { spreadsheetId, clientEmail, privateKey };
}

async function getAccessToken(clientEmail, privateKey) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  };
  const unsigned = `${base64url(JSON.stringify(header))}.${base64url(JSON.stringify(claim))}`;
  const signer = crypto.createSign("RSA-SHA256");
  signer.update(unsigned);
  signer.end();
  const signature = signer.sign(privateKey).toString("base64")
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion: `${unsigned}.${signature}`,
  });

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body,
  });
  const json = await response.json();
  if (!response.ok) {
    throw new Error(`Google token request failed: ${json.error || response.status}`);
  }
  return json.access_token;
}

function pad(row) {
  const next = [...row];
  while (next.length < COLUMN_COUNT) next.push("");
  return next.slice(0, COLUMN_COUNT);
}

function sourceLinks(labels) {
  return labels.join("\n");
}

function buildSheetValues() {
  const rows = [];
  const sectionRows = [];
  const headerRows = [];

  const add = (row = []) => {
    rows.push(pad(row));
    return rows.length - 1;
  };
  const section = (title) => {
    const idx = add([title]);
    sectionRows.push(idx);
    return idx;
  };
  const header = (row) => {
    const idx = add(row);
    headerRows.push(idx);
    return idx;
  };

  add(["REIMB_ALL_ACTIONS - Amazon reimbursement action finder"]);
  add(["Purpose", "Find cases that should be submitted now, not just past reimbursements. The current 27_REIMBURSEMENT tab remains history/dedupe."]);
  add(["Recommended shape", "One master n8n workflow as an orchestrator, with modules for inbound, AWD, ledger, returns, removals, fees, and payments. One final action queue in this tab."]);
  add(["Submission rule", "Do not auto-submit Seller Central cases. Generate case topic, evidence, draft text, and deadline; human reviews and submits."]);
  add(["Generated", new Date().toISOString().slice(0, 10), "Target tab", TAB_NAME]);
  add([]);

  section("ACTION QUEUE OUTPUT COLUMNS");
  header(["Column", "Meaning", "Example / value", "Populate from"]);
  [
    ["action_id", "Stable unique row ID for dedupe", "RWR|112-1234567-1234567|SKU1", "case_type + order/shipment/removal/event id + SKU/FNSKU"],
    ["status", "Workflow state", "NEW, REVIEW, SUBMITTED, WON, LOST, SKIP", "Default NEW"],
    ["priority", "Deadline/value priority", "P0, P1, P2", "Days left + estimated amount + confidence"],
    ["case_type", "Reimbursement family", "RWR, LEDGER_LOST, FBA_DIRECT_SHORT", "Rule module"],
    ["case_topic", "Seller Central topic/title", "FBA refund without return - order 112...", "Generated from rule template"],
    ["sku / asin / fnsku", "Product identifiers", "SKU1 / B0... / X0...", "Report rows and credential SKU map"],
    ["reference_id", "Main evidence ID", "shipment id, order id, removal order id, ledger reference id", "Rule-specific source report"],
    ["event_date", "Date that starts the clock", "refund date, return date, delivery date, ledger date", "Source report"],
    ["eligible_from", "Earliest case date", "refund date + 45/60 days; removal ship date + 15 days", "Policy window logic"],
    ["deadline", "Last safe claim date", "event date + 60/75/105/120 days", "Policy window logic"],
    ["days_left", "Urgency calculation", "8", "deadline - today"],
    ["discrepancy_qty", "Units to claim", "2", "Expected minus received/unreconciled/unreturned"],
    ["estimated_amount", "Rough claim value", "qty x manufacturing cost or proceeds", "Costs, reimbursements, settlements, fee reports"],
    ["existing_reimbursement_id", "Suppress duplicate claims", "reimbursement id if already paid", "GET_FBA_REIMBURSEMENTS_DATA"],
    ["confidence_score", "0-100 rule confidence", "85", "Evidence completeness + dedupe result"],
    ["evidence_reports", "Exactly which report rows support the case", "Ledger detail row, returns row, settlement rows", "Normalized source references"],
    ["case_text_draft", "Human-ready case body", "Please investigate...", "Generated from template"],
    ["next_step", "What the operator should do", "Submit case / wait / attach invoice / measure item", "Rule result"],
  ].forEach(add);
  add([]);

  section("REIMBURSEMENT STRATEGIES TO FIND");
  const strategyHeaderRow = header([
    "Priority",
    "Module",
    "Claim / strategy",
    "User label",
    "What it flags",
    "Amazon data to download",
    "Detection logic",
    "Dedupe / suppress if",
    "Evidence to attach",
    "Timing / urgency",
    "Action output fields",
    "Confidence",
    "Automation readiness",
    "Build phase",
    "Notes / source links",
  ]);

  const strategies = [
    ["P0", "Inbound FBA", "FBA shipment short received", "FBA DIRECT quantity mismatch", "Shipment contents show expected units greater than received units after eligible investigation date.", "Fulfillment Inbound getShipments/getShipmentItems, GET_FBA_FULFILLMENT_INBOUND_NONCOMPLIANCE_DATA, reimbursement history", "For each closed/delivered shipment line: expected_qty - received_qty - already_reimbursed_qty > 0.", "Suppress if reimbursement exists for same shipment/SKU/FNSKU/qty or shipment not eligible yet.", "Shipping plan, shipment ID, SKU/FNSKU, expected qty, received qty, POD/BOL/tracking, carton proof if available.", "Missing units: no sooner than 15 days from delivery and no later than 60 days from delivery per Amazon forum window.", "shipment_id, sku, fnsku, qty_short, delivery_date, deadline, case_topic, draft", "High when shipment status is CLOSED and discrepancy is stable.", "High", "1", sourceLinks(["Amazon inbound noncompliance report: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Claim windows: https://sellercentral.amazon.com/seller-forums/discussions/t/e62a8ace-3875-433c-8ef1-bb27a29b17c4"])],
    ["P0", "AWD/FBA transfer", "AWD to FBA replenishment short received", "FBA_FROM_AWD quantity mismatch", "AWD says inventory left/reserved for FBA replenishment, but FBA inbound/receipts show fewer units received.", "AWD API listInboundShipments/listInventory/listReplenishmentOrders, FBA inbound shipment API, FBA inventory ledger receipts", "Join AWD transfer/replenishment order to FBA shipment; flag sent_qty - fba_received_qty - reimbursed_qty > 0 after lag period.", "Suppress if later FBA receipt/found ledger event resolves the gap, or reimbursement exists.", "AWD shipment/order ID, FBA shipment ID, SKU/FNSKU, AWD sent qty, FBA received qty, dates, tracking.", "Treat as urgent; map to inbound missing unit window when the FBA shipment is delivered/checked in.", "awd_order_id, fba_shipment_id, sku, sent_qty, received_qty, qty_short, deadline", "Medium until the exact AWD to FBA reference mapping is confirmed.", "Medium", "1", sourceLinks(["AWD API tracks inbound shipments and inventory in transit to FBA: https://developer-docs.amazon.com/sp-api/lang-en_EN/docs/amazon-warehousing-and-distribution-api", "AWD inventory summaries: https://developer-docs.amazon.com/sp-api/lang-US/docs/get-inventory-summaries"])],
    ["P0", "AWD direct", "AWD direct/MCD shipment short or lost", "AWD DIRECT quantity mismatch", "AWD/MCD direct shipment expected quantity does not match delivered/received quantity at destination.", "AWD inbound/outbound shipment endpoints, AWD inventory summaries, external destination receiving file/manual upload if not in Amazon API", "For each AWD direct shipment: shipped_qty - destination_received_qty > tolerance; queue case or evidence review.", "Suppress if carrier claim already paid, destination file later reconciles, or Amazon marks rejected for seller noncompliance.", "AWD shipment ID, destination receiving report, carrier tracking/BOL/POD, SKU, quantity variance.", "Policy depends on AWD/MCD support path; use immediate review because AWD public docs do not expose a clean reimbursement window.", "awd_shipment_id, destination, sku, qty_short, evidence_status", "Medium/Low without destination receiving data.", "Medium", "2", sourceLinks(["AWD API: https://developer-docs.amazon.com/sp-api/lang-en_EN/docs/amazon-warehousing-and-distribution-api", "AWD service terms: https://supplychain.amazon.com/legal/service-terms/warehouse-and-distribution"])],
    ["P0", "Inbound FBA", "Missing shipment tracking / stale inbound shipment", "FBA SHIPMENT Missing tracking", "Shipment is expected by Amazon but missing tracking, BOL, or carrier event needed to support a claim.", "Fulfillment Inbound getShipments, shipment transport details/labels/BOL where available, carrier tracking export/manual file", "Flag shipments in WORKING/READY_TO_SHIP/SHIPPED/IN_TRANSIT/DELIVERED/CHECKED_IN beyond expected age where tracking/BOL is blank or inconsistent.", "Suppress if cancelled/deleted, seller never shipped, or tracking added later.", "Tracking number, carrier, BOL/POD, shipment labels, shipment ID, status timeline.", "This is evidence-prep, not always a reimbursement by itself; P0 when it blocks an inbound missing-unit deadline.", "shipment_id, tracking_missing_flag, status_age, evidence_needed", "High for evidence gaps; reimbursement confidence depends on quantity variance.", "High", "1", sourceLinks(["getShipments statuses and date filters: https://developer-docs.amazon.com/sp-api/reference/getshipments"])],
    ["P0", "Inbound FBA", "Inbound problem/noncompliance red alerts", "FBA SHIPMENT red stuff", "Inbound Performance report shows problem quantities, alert status, or fees that may tie to lost/damaged/short inbound units.", "GET_FBA_FULFILLMENT_INBOUND_NONCOMPLIANCE_DATA", "Flag rows with problem_quantity > 0, alert_status active, fee_total > 0, or problem_type related to quantity/label/carton/damage.", "Suppress if problem is clearly seller fault and not reimbursable; keep as coaching/ops issue if no claim.", "Inbound problem row, shipment ID, carton ID, expected/received qty, fee details.", "Immediate review; may affect 15-60 day missing-unit claims and fee disputes.", "shipment_id, problem_type, problem_qty, fee_total, recommended_case_or_ops_action", "Medium/High depending problem_type.", "High", "1", sourceLinks(["Inbound Performance report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba"])],
    ["P0", "Inventory ledger", "FC lost/misplaced unreimbursed", "REIMBURSEMENT Q/M/N/E lost and damaged", "Ledger adjustment shows lost/misplaced/damaged units and no matching reimbursement/found event.", "GET_LEDGER_DETAIL_VIEW_DATA with eventType Adjustments, GET_LEDGER_SUMMARY_VIEW_DATA, GET_FBA_REIMBURSEMENTS_DATA, FBA inventory API summaries", "Reason in M/5/lost/misplaced or damage codes; unreconciled_quantity or negative qty remains; no later Found/Reimbursed/Reconciled qty.", "Suppress if N/F found event or reimbursement covers same FNSKU/reference/qty.", "Ledger detail row, FNSKU, reason code, reference ID, quantity, reimbursement dedupe result.", "Fulfillment center lost/damaged/misplaced manual claims generally no later than 60 days after reported event.", "reference_id, reason_code, ledger_qty, unreconciled_qty, deadline, draft", "High when unreconciled and no reimbursement.", "High", "1", sourceLinks(["Ledger detail report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Reason code definitions: https://help.reasonautomation.com/seller/inventory-ledger-details", "Amazon automation/window update: https://sellercentral.amazon.com/seller-forums/discussions/t/81c3235d-4c44-47ba-96c5-883cecab3244"])],
    ["P0", "Returns/RWR", "Refund without return", "Refunds without Returns RWR", "Customer refund was issued, but no matching FBA return receipt/reimbursement after waiting period.", "Finances API listFinancialEvents, settlement V2, GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA, GET_FBA_REIMBURSEMENTS_DATA, all orders/FBA shipment sales", "Find refund event; wait required days; join to returns by order_id/SKU; if no return or reimbursement, flag.", "Suppress if return received sellable, reimbursement posted, replacement handled, or merchant-fulfilled/Safe-T path applies.", "Order ID, refund date/amount, SKU/FNSKU, no return proof, reimbursement history, settlement/finance rows.", "Customer return claim can be 60-120 days after refund/replacement per current Amazon announcement; older forum line says 45-105 for specific cases, so configure by marketplace/policy.", "order_id, refund_date, refund_amount, return_status, deadline, draft", "High when 60+ days and no return/reimbursement.", "High", "1", sourceLinks(["Returns report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "listFinancialEvents: https://developer-docs.amazon.com/sp-api/reference/listfinancialevents", "Claim windows: https://sellercentral.amazon.com/seller-forums/discussions/t/81c3235d-4c44-47ba-96c5-883cecab3244"])],
    ["P0", "Returns/RWR", "Replacement without original return", "Customer replacement original not returned", "Amazon sent replacement but original unit was not returned/reimbursed.", "GET_FBA_FULFILLMENT_CUSTOMER_SHIPMENT_REPLACEMENT_DATA, customer returns, Finances API, reimbursements", "For each replacement original_order_id: after waiting period, no return row for original and no reimbursement.", "Suppress if customer returned original or Amazon reimbursed/reversed correctly.", "Replacement order ID, original order ID, SKU, replacement reason, refund/reimbursement rows.", "Use customer return/replacement window: usually wait at least 60 days and file before 120 days from replacement/refund.", "original_order_id, replacement_order_id, sku, deadline, case_topic", "High with replacement report + no return.", "High", "1", sourceLinks(["FBA Replacements report: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba"])],
    ["P0", "Returns damage", "Returned unit damaged by Amazon/carrier", "Carrier/Amazon damaged returns", "Return report disposition is unsellable/damaged and reason indicates Amazon/carrier responsibility or no customer-fault evidence.", "GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA, ledger detail customer returns/adjustments, reimbursements", "Flag unsellable returns with detailed_disposition tied to carrier/warehouse damage; join to reimbursement.", "Suppress customer-damaged/defective/recall/policy-violation returns.", "Return report row, order ID, return reason, disposition, photos if removal returns to seller.", "Some policy notes cite carrier damaged returns within 45 days from return date; configure exact rule per marketplace.", "order_id, return_date, disposition, reason, deadline", "Medium; Amazon often disputes responsibility.", "Medium", "2", sourceLinks(["Customer returns report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Window detail: https://sellercentral.amazon.com/seller-forums/discussions/t/e62a8ace-3875-433c-8ef1-bb27a29b17c4"])],
    ["P0", "Removals", "Removal order lost in transit", "Removal order missing/lost", "Removal order detail says shipped/in-process but shipment detail/tracking/delivery never proves seller received inventory.", "GET_FBA_FULFILLMENT_REMOVAL_ORDER_DETAIL_DATA, GET_FBA_FULFILLMENT_REMOVAL_SHIPMENT_DETAIL_DATA, carrier tracking/manual receiving file, reimbursements", "requested_qty - received_by_seller_qty - reimbursed_qty > 0; or tracking never delivered after ship date.", "Suppress if disposed/cancelled, seller confirmed receipt, or reimbursement exists.", "Removal order ID, tracking, shipped qty, receiving count, box photos if received short.", "Lost removals: no sooner than 15 and no later than 75 days from shipment creation/ship date.", "removal_order_id, tracking, qty_missing, ship_date, deadline", "High when tracking not delivered or received short.", "High", "1", sourceLinks(["Removal order/shipment report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Removal windows: https://sellercentral.amazon.com/seller-forums/discussions/t/81c3235d-4c44-47ba-96c5-883cecab3244"])],
    ["P0", "Removals", "Removal order damaged/different/incomplete after sent good", "Removal order damaged units received", "Seller receives damaged, different, or incomplete units from a removal shipment that were expected/sent as sellable/good.", "Removal shipment detail, removal order detail, seller receiving/photo upload, ledger disposition before removal", "If removal disposition was sellable/good or not seller-damaged, and seller receiving marks damaged/different/incomplete, flag with photos required.", "Suppress if removal disposition already unsellable/customer damaged/defective or no receiving evidence.", "Photos all sides, packaging, removal ID, SKU/FNSKU, shipment tracking, before-removal disposition.", "Damaged/different/incomplete removal received: within 30 days from delivery date per Amazon forum window.", "removal_order_id, sku, condition_received, photo_required, deadline", "Medium/High with photos.", "Medium", "1", sourceLinks(["Removal reports: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Window detail: https://sellercentral.amazon.com/seller-forums/discussions/t/e62a8ace-3875-433c-8ef1-bb27a29b17c4"])],
    ["P1", "Fee/Cubiscan", "FBA size tier or dimension overcharge", "Cubicscan recalculate fba size item", "Amazon-measured package dimensions/weight push ASIN into a higher fee/storage tier than expected.", "GET_FBA_ESTIMATED_FBA_FEES_TXT_DATA, GET_FBA_STORAGE_FEE_CHARGES_DATA, Catalog Items dimensions, product fee estimate API, seller measurements/photos", "Compare Amazon longest/median/shortest side, weight, size tier, expected fee vs known package specs or prior fee tier; flag changes and overcharge deltas.", "Suppress if current measured dimensions match seller evidence or fee tier change is legitimate packaging change.", "Photos with ruler/scale, ASIN/SKU, Amazon fee preview, storage fee volume, previous fee if available.", "File fee/dimension dispute quickly; many reimbursement windows are short and operationally treated as 60 days.", "asin, sku, amazon_dims, seller_dims, expected_fee, charged_fee, delta", "Medium; requires physical measurement proof.", "Medium", "2", sourceLinks(["FBA Fee Preview fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Storage fee dimensions/volume fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba"])],
    ["P1", "Storage/Aged fees", "Aged inventory surcharge overcharge", "Aged inventory surcharge error", "Aged inventory surcharge or long-term storage fee does not match age buckets, volume, or quantity charged.", "GET_FBA_FULFILLMENT_LONGTERM_STORAGE_FEE_CHARGES_DATA, GET_FBA_INVENTORY_PLANNING_DATA, GET_FBA_STORAGE_FEE_CHARGES_DATA, settlement V2", "Recalculate amount_charged = rate_surcharge x qty_charged and compare volume/age bucket against inventory planning/aged data.", "Suppress if amount matches Amazon report and no age/qty/volume dispute.", "LTSF report row, inventory planning age bucket, storage fee report volume, settlement fee row.", "Review monthly as soon as charges post; exact dispute window should be configured per current Seller Central policy.", "sku, fnsku, age_tier, qty_charged, charged_amount, expected_amount, delta", "Medium", "Medium", "2", sourceLinks(["Long-term storage fee fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Inventory planning aged fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba"])],
    ["P1", "Storage fees", "Monthly storage/overage fee error", "Storage fee / overage fee error", "Storage fee amount, cubic volume, or overage fee is inconsistent with inventory quantities and product volume.", "GET_FBA_STORAGE_FEE_CHARGES_DATA, GET_FBA_OVERAGE_FEE_CHARGES_DATA, inventory planning, settlement V2", "Recalculate estimated_total_item_volume and fee; flag large delta, sudden volume jump, wrong storage type, or overage mismatch.", "Suppress known seasonal rate changes and valid dangerous goods classification.", "Storage fee row, overage fee row, dimensions, qty on hand, settlement charge.", "Monthly review; escalate before evidence ages out.", "asin, fnsku, storage_type, volume, charged_fee, expected_fee, delta", "Medium", "Medium", "2", sourceLinks(["Storage fee report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Overage fee report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba"])],
    ["P1", "Payments", "Reimbursement not paid or reversed unexpectedly", "Payment / settlement / reimbursement amount cases", "Reimbursement report says approved, but settlement/finances does not show the payment; or a reversal appears without matching found unit.", "GET_FBA_REIMBURSEMENTS_DATA, GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2, Finances API, ledger detail", "Join reimbursement_id/case/order/SKU to settlement amount type/description; flag approved amount not paid, underpaid, or reversed without found/reconciled event.", "Suppress if settlement not generated yet or deferred event unreleased.", "Reimbursement row, settlement rows, finance event, ledger found/reversal evidence.", "Dispute valuation within 60 days after reimbursement per Amazon policy PDF; missing payout review as soon as settlement closes.", "reimbursement_id, case_id, amount_expected, amount_paid, delta, settlement_id", "High after settlement closes.", "High", "1", sourceLinks(["Reimbursements report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba", "Settlement V2 fields: https://developer-docs.amazon.com/sp-api/lang-en_EN/docs/report-type-values-settlement", "Valuation dispute window: https://m.media-amazon.com/images/G/01/rainier/help./PC__Kiwi_Comms_redlined_policy_US_clean.pdf"])],
    ["P1", "Payments", "Underpaid reimbursement / valuation dispute", "Manufacturing cost / amount dispute", "Amazon reimbursed, but amount is below approved manufacturing cost or expected proceeds.", "GET_FBA_REIMBURSEMENTS_DATA, settlement V2, seller manufacturing cost table/manual input, order sales history", "Compare amount_per_unit to stored manufacturing cost or expected proceeds; flag materially low values.", "Suppress if item unsellable/discontinued and policy supports reduced value; suppress if already disputed in last 30 days with no new proof.", "Invoice/cost proof, reimbursement ID, amount paid, expected manufacturing cost/proceeds.", "Amazon policy PDF says valuation disputes can be filed within 60 days after reimbursement.", "sku, reimbursement_id, amount_per_unit, expected_value, delta, deadline", "Medium/High with cost documentation.", "Medium", "1", sourceLinks(["Reimbursement valuation policy: https://m.media-amazon.com/images/G/01/rainier/help./PC__Kiwi_Comms_redlined_policy_US_clean.pdf"])],
    ["P1", "Refund accuracy", "Customer refund overage / inaccurate order refund", "Refund overage / fee mismatch", "Refund amount or fees charged to seller exceed original order economics or expected fee refund.", "Finances API order events, settlement V2, all orders/FBA shipment sales, returns report", "For each refund: compare refund item/shipping/tax/fee amounts to original order and return policy; flag refund > paid or fee credit missing.", "Suppress promotions/taxes/marketplace-withheld cases that explain delta.", "Order financial events, original order sale, refund event, settlement rows.", "Review weekly; exact path may be reimbursement or Safe-T depending fulfillment and marketplace.", "order_id, sku, refund_total, original_total, fee_delta, action", "Medium", "Medium", "2", sourceLinks(["Finances API: https://developer-docs.amazon.com/sp-api/reference/listfinancialevents", "Settlement V2: https://developer-docs.amazon.com/sp-api/lang-en_EN/docs/report-type-values-settlement"])],
    ["P1", "Inventory ledger", "Wrong account transfer / redefinition loss", "Inventory adjustment transfer mismatch", "Ledger transfer-out/product redefinition removes inventory without matching transfer-in or sale/removal/reimbursement.", "GET_LEDGER_DETAIL_VIEW_DATA, inventory ledger reason codes, inventory summaries", "Pair code 4 transfer-out with code 3 transfer-in or equivalent; flag unmatched negative movements.", "Suppress if net by ASIN/FNSKU is zero after mapping old/new identifiers.", "Ledger references, old/new SKU/FNSKU, quantities, title mapping.", "Treat as 60-day inventory adjustment review.", "reference_id, old_fnsku, new_fnsku, qty_unmatched", "Medium", "Medium", "3", sourceLinks(["Ledger fields and codes: https://help.reasonautomation.com/seller/inventory-ledger-details"])],
    ["P2", "Disposal/liquidation", "Unauthorized disposal/liquidation discrepancy", "Destroyed/disposed items", "Ledger/removal report shows disposal or liquidation not requested/expected and no reimbursement.", "GET_LEDGER_DETAIL_VIEW_DATA, removal order detail, reimbursements, settlement V2", "Flag disposed/liquidated quantities with no matching seller removal request, reimbursement, or policy reason.", "Suppress seller-requested disposal/removal or expired/restricted units where Amazon policy excludes reimbursement.", "Ledger disposal row, removal order status, SKU/FNSKU, policy reason.", "Manual review; reimbursability depends on why Amazon disposed/liquidated.", "sku, fnsku, disposal_qty, reference_id, action", "Low/Medium", "Medium", "3", sourceLinks(["GETIDA and Refunds Manager both list disposed/destroyed items as recoverable categories: https://getida.com/ | https://www.refundsmanager.com/"])],
    ["P2", "Stranded/inventory", "Stranded inventory caused by Amazon issue", "Stranded / listing blocked leakage", "Inventory stranded due to Amazon system/catalog issue and storage fees continue or inventory becomes unfulfillable.", "GET_STRANDED_INVENTORY_UI_DATA, inventory planning, storage fee reports", "Flag stranded rows with Amazon-caused status/error and high fees/days stranded.", "Suppress seller listing errors that require catalog fixes, not reimbursement.", "Stranded report row, error message, storage fee impact, case history.", "Operational case first; reimbursement only if Amazon error caused measurable fee/loss.", "sku, fnsku, stranded_reason, fees_estimate, next_step", "Low/Medium", "Medium", "3", sourceLinks(["Stranded inventory report fields: https://developer-docs.amazon.com/sp-api/lang-en_US/docs/report-type-values-fba"])],
    ["P2", "AWD fees", "AWD storage/processing/transport fee discrepancy", "AWD fee discrepancy", "AWD fees in Seller Central/payment reports do not match AWD inventory/shipment activity.", "AWD API inventory/shipments, Seller Central AWD fee exports/manual, settlement V2/Finances", "Recalculate fee drivers where data exists; flag fee without shipment/inventory basis or large unexplained jumps.", "Suppress if fee schedule/usage supports charge.", "AWD fee export, settlement row, AWD inventory and shipment rows.", "Manual review because public AWD docs say fee info is in Seller Central, not fully exposed via API.", "fee_type, charged_amount, expected_amount, evidence_gap", "Low until fee source is automated.", "Low", "3", sourceLinks(["AWD docs note fee info in Seller Central: https://developer-docs.amazon.com/sp-api/lang-en_EN/docs/amazon-warehousing-and-distribution-api"])],
  ];
  strategies.forEach(add);
  add([]);

  section("REPORT DOWNLOAD MAP FOR THE WORKFLOW");
  header(["Data source", "Amazon report/API", "Why it is needed", "Primary strategies", "Frequency", "Notes"]);
  [
    ["Inventory ledger detail", "GET_LEDGER_DETAIL_VIEW_DATA", "Core inventory movement proof: lost, damaged, found, disposed, receipts, returns.", "Q/M/N/E lost/damaged, transfer mismatch, disposal", "Daily", "Use eventType Adjustments first; pull full detail for references."],
    ["Inventory ledger summary", "GET_LEDGER_SUMMARY_VIEW_DATA", "Monthly/daily reconciliation by SKU/FNSKU/FC/country.", "Ledger sanity checks", "Daily/weekly", "Useful for aggregate validation."],
    ["Reimbursements", "GET_FBA_REIMBURSEMENTS_DATA", "Dedupe, payment amount, case ID, reason, reversal linkage.", "All claim types", "Daily", "This is history, not action discovery by itself."],
    ["Inbound shipments", "Fulfillment Inbound getShipments + shipment items", "Shipment status, expected/received quantities, delivery timing.", "FBA direct, missing tracking", "Daily", "Needed for actual short-shipment queue."],
    ["Inbound performance", "GET_FBA_FULFILLMENT_INBOUND_NONCOMPLIANCE_DATA", "Problem type, expected/received qty, alert status, inbound fees.", "FBA red stuff, inbound disputes", "Daily", "Good early-warning feed."],
    ["AWD inventory/shipments", "AWD listInventory/listInboundShipments/listReplenishmentOrders", "AWD quantities and shipments in transit to AWD/FBA.", "FBA_FROM_AWD, AWD_DIRECT", "Daily", "US marketplace only per docs; fee data may require Seller Central/manual export."],
    ["Customer returns", "GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA", "Return received status, reason, disposition, comments.", "RWR, damaged returns", "Daily", "Join to refunds and replacements."],
    ["Replacements", "GET_FBA_FULFILLMENT_CUSTOMER_SHIPMENT_REPLACEMENT_DATA", "Original/replacement order mapping.", "Replacement without return", "Daily", "NA and IN marketplace availability per docs."],
    ["Finances", "listFinancialEvents", "Order/refund/reimbursement financial events before settlement closes.", "RWR, refund overage, payment reconciliation", "Daily", "Events from last 48 hours may be delayed."],
    ["Settlement V2", "GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2", "Final payment posting proof and fee/refund/reimbursement amounts.", "Payment mismatch, fee disputes", "Per settlement", "Cannot be requested; search generated reports."],
    ["Removal order detail", "GET_FBA_FULFILLMENT_REMOVAL_ORDER_DETAIL_DATA", "Requested/cancelled/disposed/shipped/in-process removal quantities.", "Removal lost/damaged/incomplete", "Daily", "Near real-time."],
    ["Removal shipment detail", "GET_FBA_FULFILLMENT_REMOVAL_SHIPMENT_DETAIL_DATA", "Carrier, tracking number, shipped quantity for removals.", "Removal lost/damaged/incomplete", "Daily", "Does not include canceled returns or disposed items."],
    ["Fee preview", "GET_FBA_ESTIMATED_FBA_FEES_TXT_DATA", "Amazon dimensions, weight, size tier, estimated fees.", "Cubiscan, FBA fee overcharge", "Daily max", "Docs caution once per day per seller."],
    ["Storage fee report", "GET_FBA_STORAGE_FEE_CHARGES_DATA", "Monthly storage dimensions, volume, rate, fee.", "Cubiscan, storage errors", "Monthly", "Can be requested/scheduled."],
    ["Long-term storage fee", "GET_FBA_FULFILLMENT_LONGTERM_STORAGE_FEE_CHARGES_DATA", "Aged inventory surcharge/long-term storage charge details.", "Aged inventory surcharge", "Monthly", "Request exactly one month per docs."],
    ["Inventory planning", "GET_FBA_INVENTORY_PLANNING_DATA", "Aged/excess units, estimated fees, inbound and inventory health.", "Aged fee, storage, inbound cross-check", "Daily/weekly", "Good validation source."],
    ["FBA inventory summary", "FBA Inventory API getInventorySummaries", "Fulfillable/inbound/reserved/unfulfillable/researching quantities.", "Current-state reconciliation", "Daily", "Useful for resolving false positives."],
  ].forEach(add);
  add([]);

  section("COMPETITOR / APP PATTERN TO MIRROR");
  header(["Service", "Publicly visible approach", "Claim families they mention", "What to copy into our workflow", "Source"]);
  [
    ["GETIDA", "Continuous FBA transaction audits, discrepancy detection, claim management by specialists.", "Lost/damaged, overcharges, inbound discrepancies, removals, fee issues.", "Continuous audit + case tracking + reimbursement history dedupe.", "https://getida.com/"],
    ["Refunds Manager", "Imports FBA transactions, uses software plus auditor review, submits eligible cases manually.", "Missing reimbursements, replacements, refunds, destroyed items, warehouse losses, inventory adjustments, shipment stock counts, damaged returns, weight/dimension fees, commission fees, missing removal orders.", "Do not auto-send; generate clean evidence and manual review queue.", "https://www.refundsmanager.com/"],
    ["Helium 10 Managed Refund Service", "End-to-end claim specialists audit FBA transactions, identify refund opportunities, submit and follow up.", "Inbound, warehouse, removals, dimensions, customer returns/exchanges.", "Separate modules by claim family; track case lifecycle.", "https://www.helium10.com/tools/operations/managed-refund-service/"],
    ["SPS Revenue Recovery 3P", "Automated audit and managed claim support with dashboard/reporting.", "Shipment discrepancies, lost/damaged inventory, fee errors, other FBA categories.", "Action dashboard, deadlines, exportable detail logs.", "https://www.spscommerce.com/products/revenue-recovery/3p/"],
  ].forEach(add);
  add([]);

  section("N8N ARCHITECTURE RECOMMENDATION");
  header(["Layer", "Node/module", "What it does", "Reason"]);
  [
    ["1", "Trigger / manual run", "Run daily, plus manual webhook run for testing.", "Claim windows are short; daily is safer."],
    ["2", "Credentials / date window", "Use existing dynamic credentials and configurable lookback windows.", "Keeps current workflow style."],
    ["3", "Download modules", "One module each for inbound, AWD, ledger, returns, removals, fees, payments.", "Keeps the master workflow understandable."],
    ["4", "Normalize rows", "Convert every Amazon row into common IDs: SKU, ASIN, FNSKU, order/shipment/removal/reference, date, qty, amount.", "Makes cross-report joins possible."],
    ["5", "Rule engine", "Run strategy rules and produce candidate actions.", "The business logic lives here."],
    ["6", "Dedupe engine", "Compare candidates against reimbursements, settlements, prior action IDs, found events, and status.", "Prevents duplicate/premature cases."],
    ["7", "Scoring/deadlines", "Calculate confidence, eligible_from, deadline, days_left, priority, evidence completeness.", "This is what makes the sheet actionable."],
    ["8", "Sheet writer", "Append/upsert into REIMB_ALL_ACTIONS; preserve user status/notes.", "Humans can work the queue."],
    ["9", "No case upload", "Generate topic and case text only.", "Safer and cleaner with Amazon policy; final submission stays human-reviewed."],
  ].forEach(add);

  return { rows, sectionRows, headerRows, strategyHeaderRow };
}

async function googleRequest(token, url, options = {}) {
  let lastError;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const response = await fetch(url, {
      ...options,
      headers: {
        authorization: `Bearer ${token}`,
        "content-type": "application/json",
        ...(options.headers || {}),
      },
    });
    const text = await response.text();
    let json = {};
    if (text) {
      try {
        json = JSON.parse(text);
      } catch {
        json = { raw: text };
      }
    }
    if (response.ok) return json;
    lastError = new Error(`Google Sheets API failed ${response.status} for ${url.replace(/\?.*$/, "")}: ${JSON.stringify(json).slice(0, 500)}`);
    if (response.status >= 500 && attempt < 3) {
      await new Promise((resolve) => setTimeout(resolve, 1000 * (attempt + 1)));
      continue;
    }
    throw lastError;
  }
  throw lastError;
}

function isMissingSheetError(error) {
  return /Unable to parse range|not found|Cannot find|Unable to find|No grid with id/i.test(error.message);
}

async function createSheet(token, base) {
  const createResult = await googleRequest(token, `${base}:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      requests: [{ addSheet: { properties: { title: TAB_NAME, index: 0 } } }],
    }),
  });
  return createResult.replies?.[0]?.addSheet?.properties?.sheetId || null;
}

async function main() {
  const { spreadsheetId, clientEmail, privateKey } = await loadCredentials();
  const token = await getAccessToken(clientEmail, privateKey);
  const base = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;

  const { rows, sectionRows, headerRows, strategyHeaderRow } = buildSheetValues();
  const normalizedRows = rows.map(pad);

  let sheetId = null;
  try {
    const metadata = await googleRequest(token, `${base}?fields=sheets(properties(sheetId,title,index))`);
    const sheet = metadata.sheets?.find((s) => s.properties.title === TAB_NAME);
    sheetId = sheet?.properties?.sheetId || null;
    if (sheetId === null) sheetId = await createSheet(token, base);
  } catch (error) {
    console.warn(`Metadata lookup skipped: ${error.message}`);
  }

  try {
    await googleRequest(token, `${base}/values/${encodeURIComponent(`${TAB_NAME}!A:O`)}:clear`, {
      method: "POST",
      body: JSON.stringify({}),
    });
  } catch (error) {
    if (!isMissingSheetError(error)) throw error;
    sheetId = await createSheet(token, base);
  }

  await googleRequest(token, `${base}/values/${encodeURIComponent(`${TAB_NAME}!A1:O${normalizedRows.length}`)}?valueInputOption=RAW`, {
    method: "PUT",
    body: JSON.stringify({ values: normalizedRows }),
  });

  if (sheetId === null) {
    try {
      const metadata = await googleRequest(token, `${base}?fields=sheets(properties(sheetId,title,index))`);
      const sheet = metadata.sheets?.find((s) => s.properties.title === TAB_NAME);
      sheetId = sheet?.properties?.sheetId || null;
    } catch (error) {
      console.warn(`Formatting skipped because sheet metadata is unavailable: ${error.message}`);
    }
  }

  if (sheetId === null) {
    const verifyOnly = await googleRequest(
      token,
      `${base}/values/${encodeURIComponent(`${TAB_NAME}!A1:O${Math.min(normalizedRows.length, 80)}`)}?majorDimension=ROWS`
    );
    console.log(JSON.stringify({
      spreadsheetId,
      tab: TAB_NAME,
      rowsWritten: normalizedRows.length,
      columnsWritten: COLUMN_COUNT,
      verifiedRows: verifyOnly.values?.length || 0,
      formatted: false,
    }, null, 2));
    return;
  }

  const headerRequests = headerRows.map((rowIndex) => ({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: 0,
        endColumnIndex: COLUMN_COUNT,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: { red: 0.88, green: 0.93, blue: 0.98 },
          textFormat: { bold: true, foregroundColor: { red: 0.08, green: 0.16, blue: 0.25 } },
          wrapStrategy: "WRAP",
          horizontalAlignment: "LEFT",
          verticalAlignment: "MIDDLE",
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat,wrapStrategy,horizontalAlignment,verticalAlignment)",
    },
  }));

  const sectionFormatRequests = sectionRows.map((rowIndex) => ({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: 0,
        endColumnIndex: COLUMN_COUNT,
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: { red: 0.11, green: 0.23, blue: 0.34 },
          textFormat: { bold: true, fontSize: 12, foregroundColor: { red: 1, green: 1, blue: 1 } },
          wrapStrategy: "WRAP",
        },
      },
      fields: "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)",
    },
  }));

  const columnWidths = [90, 140, 220, 180, 300, 310, 360, 280, 300, 260, 260, 150, 160, 110, 360];
  const widthRequests = columnWidths.map((pixelSize, index) => ({
    updateDimensionProperties: {
      range: { sheetId, dimension: "COLUMNS", startIndex: index, endIndex: index + 1 },
      properties: { pixelSize },
      fields: "pixelSize",
    },
  }));

  const formatRequests = [
    {
      repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COLUMN_COUNT },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.05, green: 0.18, blue: 0.29 },
            textFormat: { bold: true, fontSize: 16, foregroundColor: { red: 1, green: 1, blue: 1 } },
            horizontalAlignment: "LEFT",
          },
        },
        fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
      },
    },
    ...sectionFormatRequests,
    ...headerRequests,
    {
      repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: normalizedRows.length, startColumnIndex: 0, endColumnIndex: COLUMN_COUNT },
        cell: {
          userEnteredFormat: {
            wrapStrategy: "WRAP",
            verticalAlignment: "TOP",
          },
        },
        fields: "userEnteredFormat(wrapStrategy,verticalAlignment)",
      },
    },
    {
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { frozenRowCount: Math.min(strategyHeaderRow + 1, normalizedRows.length) },
        },
        fields: "gridProperties.frozenRowCount",
      },
    },
    {
      updateDimensionProperties: {
        range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: normalizedRows.length },
        properties: { pixelSize: 44 },
        fields: "pixelSize",
      },
    },
    {
      updateDimensionProperties: {
        range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: 1 },
        properties: { pixelSize: 34 },
        fields: "pixelSize",
      },
    },
    ...widthRequests,
  ];

  let formatted = false;
  try {
    await googleRequest(token, `${base}:batchUpdate`, {
      method: "POST",
      body: JSON.stringify({ requests: formatRequests }),
    });
    formatted = true;
  } catch (error) {
    console.warn(`Formatting skipped: ${error.message}`);
  }

  const verify = await googleRequest(
    token,
    `${base}/values/${encodeURIComponent(`${TAB_NAME}!A1:O${Math.min(normalizedRows.length, 80)}`)}?majorDimension=ROWS`
  );

  console.log(JSON.stringify({
    spreadsheetId,
    tab: TAB_NAME,
    rowsWritten: normalizedRows.length,
    columnsWritten: COLUMN_COUNT,
    verifiedRows: verify.values?.length || 0,
    formatted,
  }, null, 2));
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
