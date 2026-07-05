import { createWriteStream } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, join } from "node:path";
import { pipeline } from "node:stream/promises";

const API_BASE = "https://api.midapi.ai/api/v1";
const CREDIT_COST = Number(process.env.MIDAPI_CREDIT_COST || 3);
const POLL_INTERVAL_MS = Number(process.env.MIDAPI_POLL_INTERVAL_MS || 30000);
const MAX_POLL_MINUTES = Number(process.env.MIDAPI_MAX_POLL_MINUTES || 45);
const SUBMIT_DELAY_MS = Number(process.env.MIDAPI_SUBMIT_DELAY_MS || 1200);
const PROMPT =
  process.env.MIDAPI_PROMPT ||
  "make art https://ibb.co/PvJR6pYm --ow 1000 --cref https://ibb.co/PvJR6pYm --sref https://ibb.co/PvJR6pYm --v 7.0 --ar 3:4";

const key = process.env.MIDAPI_KEY;
if (!key) {
  throw new Error("Set MIDAPI_KEY in the environment.");
}

const now = new Date();
const runId = now.toISOString().replace(/[:.]/g, "-");
const outDir = process.env.MIDAPI_OUT_DIR || join(process.cwd(), "midapi_runs", runId);
const rawDir = join(outDir, "raw");
const imageDir = join(outDir, "images");

async function sleep(ms) {
  await new Promise((resolve) => setTimeout(resolve, ms));
}

async function api(path, options = {}) {
  const headers = {
    Authorization: `Bearer ${key}`,
    "Content-Type": "application/json",
    ...(options.headers || {}),
  };
  for (const [name, value] of Object.entries(headers)) {
    if (value === undefined || value === null) delete headers[name];
  }
  const res = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers,
  });
  const text = await res.text();
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    json = { raw: text };
  }
  if (!res.ok) {
    throw new Error(`${res.status} ${res.statusText}: ${text}`);
  }
  return json;
}

async function getCredits() {
  const result = await api("/common/credit", { method: "GET" });
  if (result.code !== 200) {
    throw new Error(`Credit check failed: ${JSON.stringify(result)}`);
  }
  return Number(result.data);
}

async function submitJob(index) {
  const body = {
    taskType: "mj_txt2img",
    speed: "relaxed",
    prompt: PROMPT,
    enableTranslation: false,
  };
  const result = await api("/mj/generate", {
    method: "POST",
    body: JSON.stringify(body),
  });
  await writeFile(join(rawDir, `submit-${String(index).padStart(3, "0")}.json`), JSON.stringify(result, null, 2));
  if (result.code !== 200 || !result.data?.taskId) {
    throw new Error(`Submit failed: ${JSON.stringify(result)}`);
  }
  return result.data.taskId;
}

async function fetchJob(taskId) {
  const result = await api(`/mj/record-info?taskId=${encodeURIComponent(taskId)}`, {
    method: "GET",
    headers: { "Content-Type": undefined },
  });
  await writeFile(join(rawDir, `task-${taskId}.json`), JSON.stringify(result, null, 2));
  return result;
}

function parseResultInfo(value) {
  if (!value) return null;
  if (typeof value === "object") return value;
  try {
    return JSON.parse(value);
  } catch {
    return { raw: value };
  }
}

function collectUrls(value, urls = new Set()) {
  if (!value) return urls;
  if (typeof value === "string") {
    if (/^https?:\/\/.+\.(png|jpe?g|webp)(\?.*)?$/i.test(value)) urls.add(value);
    return urls;
  }
  if (Array.isArray(value)) {
    for (const item of value) collectUrls(item, urls);
    return urls;
  }
  if (typeof value === "object") {
    for (const item of Object.values(value)) collectUrls(item, urls);
  }
  return urls;
}

async function downloadUrl(url, index) {
  const parsed = new URL(url);
  const cleanName = basename(parsed.pathname) || `image-${index}.png`;
  const filename = `${String(index).padStart(3, "0")}-${cleanName}`;
  const path = join(imageDir, filename);
  const res = await fetch(url);
  if (!res.ok || !res.body) {
    throw new Error(`Download failed for ${url}: ${res.status} ${res.statusText}`);
  }
  await pipeline(res.body, createWriteStream(path));
  return path;
}

await mkdir(rawDir, { recursive: true });
await mkdir(imageDir, { recursive: true });

const startingCredits = await getCredits();
const requestedJobs = Math.floor(startingCredits / CREDIT_COST);
const maxJobs = Number(process.env.MIDAPI_MAX_JOBS || requestedJobs);
const jobCount = Math.min(requestedJobs, maxJobs);

const manifest = {
  runId,
  outDir,
  prompt: PROMPT,
  mode: "relaxed",
  taskType: "mj_txt2img",
  creditCost: CREDIT_COST,
  startingCredits,
  plannedJobs: jobCount,
  tasks: [],
  images: [],
  errors: [],
};

await writeFile(join(outDir, "manifest.json"), JSON.stringify(manifest, null, 2));
console.log(`Run folder: ${outDir}`);
console.log(`Starting credits: ${startingCredits}`);
console.log(`Submitting ${jobCount} Relax jobs at ${CREDIT_COST} credits each.`);

for (let i = 1; i <= jobCount; i++) {
  try {
    const taskId = await submitJob(i);
    manifest.tasks.push({ index: i, taskId, status: "submitted" });
    console.log(`Submitted ${i}/${jobCount}: ${taskId}`);
  } catch (error) {
    manifest.errors.push({ stage: "submit", index: i, message: error.message });
    console.error(`Submit ${i} failed: ${error.message}`);
  }
  await writeFile(join(outDir, "manifest.json"), JSON.stringify(manifest, null, 2));
  if (i < jobCount) await sleep(SUBMIT_DELAY_MS);
}

const deadline = Date.now() + MAX_POLL_MINUTES * 60 * 1000;
let unfinished = manifest.tasks.filter((task) => task.status === "submitted" || task.status === "running");

while (unfinished.length > 0 && Date.now() < deadline) {
  for (const task of unfinished) {
    try {
      const result = await fetchJob(task.taskId);
      const data = result.data || {};
      if (data.successFlag === 1 || data.successFlag === 2) {
        task.status = "success";
        task.completeTime = data.completeTime;
        task.resultInfoJson = parseResultInfo(data.resultInfoJson);
        const urls = [...collectUrls(task.resultInfoJson)];
        task.imageUrls = urls;
        console.log(`Completed ${task.index}: ${task.taskId}, ${urls.length} image URLs`);
      } else if (data.successFlag === 3 || data.errorCode || data.errorMessage) {
        task.status = "failed";
        task.completeTime = data.completeTime;
        task.errorCode = data.errorCode;
        task.errorMessage = data.errorMessage;
        console.log(`Failed ${task.index}: ${task.taskId}: ${data.errorMessage || data.errorCode || "unknown error"}`);
      } else {
        task.status = "running";
      }
    } catch (error) {
      manifest.errors.push({ stage: "poll", taskId: task.taskId, message: error.message });
      console.error(`Poll failed for ${task.taskId}: ${error.message}`);
    }
  }
  await writeFile(join(outDir, "manifest.json"), JSON.stringify(manifest, null, 2));
  unfinished = manifest.tasks.filter((task) => task.status === "submitted" || task.status === "running");
  if (unfinished.length > 0) {
    console.log(`${unfinished.length} jobs still running; sleeping ${Math.round(POLL_INTERVAL_MS / 1000)}s.`);
    await sleep(POLL_INTERVAL_MS);
  }
}

let downloadIndex = 1;
const downloaded = new Set();
for (const task of manifest.tasks.filter((item) => item.status === "success")) {
  for (const url of task.imageUrls || []) {
    if (downloaded.has(url)) continue;
    try {
      const path = await downloadUrl(url, downloadIndex++);
      downloaded.add(url);
      manifest.images.push({ taskId: task.taskId, url, path });
      console.log(`Downloaded ${path}`);
    } catch (error) {
      manifest.errors.push({ stage: "download", taskId: task.taskId, url, message: error.message });
      console.error(`Download failed: ${error.message}`);
    }
  }
}

manifest.endingCredits = await getCredits();
manifest.finishedAt = new Date().toISOString();
await writeFile(join(outDir, "manifest.json"), JSON.stringify(manifest, null, 2));
await writeFile(join(outDir, "summary.txt"), [
  `Run folder: ${outDir}`,
  `Prompt: ${PROMPT}`,
  `Starting credits: ${manifest.startingCredits}`,
  `Ending credits: ${manifest.endingCredits}`,
  `Planned jobs: ${manifest.plannedJobs}`,
  `Submitted jobs: ${manifest.tasks.length}`,
  `Successful jobs: ${manifest.tasks.filter((task) => task.status === "success").length}`,
  `Failed jobs: ${manifest.tasks.filter((task) => task.status === "failed").length}`,
  `Downloaded files: ${manifest.images.length}`,
  `Errors: ${manifest.errors.length}`,
  "",
].join("\n"));

console.log("Done.");
console.log(`Ending credits: ${manifest.endingCredits}`);
console.log(`Downloaded files: ${manifest.images.length}`);
