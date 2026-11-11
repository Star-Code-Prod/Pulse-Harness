// Tool support that needs more room than a one-liner: git, test runners,
// document round-trips, memory, templates, charts, schedules and traces.
// Everything takes the project root so the tools stay thin.

import fs from "node:fs/promises";
import fssync from "node:fs";
import path from "node:path";
import os from "node:os";
import { execFile } from "node:child_process";
import { extractOffice } from "./office.js";
import { writeDocx, writeXlsx, writePptx } from "./docwrite.js";

const STAR_DIR = ".star-code";
const star = (root, ...parts) => path.join(root, STAR_DIR, ...parts);
const rel = (root, p) => path.relative(root, p) || p;
const ext = (p) => path.extname(String(p || "")).toLowerCase();

function run(bin, args, root, timeout = 60_000) {
  return new Promise((resolve) => {
    execFile(bin, args, { cwd: root, timeout, windowsHide: true, maxBuffer: 8 * 1024 * 1024 }, (err, stdout, stderr) => {
      resolve({
        ok: !err,
        code: err && typeof err.code === "number" ? err.code : err ? 1 : 0,
        out: String(stdout || ""),
        err: String(stderr || ""),
      });
    });
  });
}

async function readJson(file, fallback) {
  try {
    return JSON.parse(await fs.readFile(file, "utf8"));
  } catch {
    return fallback;
  }
}

async function writeJson(file, value) {
  await fs.mkdir(path.dirname(file), { recursive: true });
  await fs.writeFile(file, JSON.stringify(value, null, 2) + "\n", "utf8");
}

/* -------------------------------- git --------------------------------- */

function gitMissing(r) {
  return "git is not installed or not on PATH - install Git for Windows (or Xcode tools on macOS) and try again.";
}

export async function gitStatus(root) {
  const branch = await run("git", ["rev-parse", "--abbrev-ref", "HEAD"], root, 15_000);
  if (!branch.ok && /not a git repository/i.test(branch.err)) return "This folder is not a git repository. Run git_commit after `git init`, or work without version control.";
  if (!branch.ok && /ENOENT|not recognized/i.test(branch.err || branch.out)) return gitMissing(root);
  const files = await run("git", ["status", "--porcelain=v1", "--branch"], root, 20_000);
  const lines = files.out.split("\n").filter(Boolean);
  const head = lines.find((l) => l.startsWith("##")) || "## " + branch.out.trim();
  const changes = lines.filter((l) => !l.startsWith("##"));
  const staged = changes.filter((l) => /^[MADRC]/.test(l)).length;
  const unstaged = changes.filter((l) => /^.[MADRC]/.test(l)).length;
  const untracked = changes.filter((l) => l.startsWith("??")).length;
  const out = [head.trim()];
  if (!changes.length) out.push("clean - nothing to commit");
  else {
    out.push(staged + " staged, " + unstaged + " modified, " + untracked + " untracked");
    out.push(...changes.slice(0, 200));
    if (changes.length > 200) out.push("[capped at 200 files]");
  }
  return out.join("\n");
}

export async function gitDiff(root, { path: p = "", staged = false, stat = false } = {}) {
  const args = ["diff", "--no-color"];
  if (stat) args.push("--stat");
  if (staged) args.push("--cached");
  if (p) args.push("--", p);
  const res = await run("git", args, root, 30_000);
  if (!res.ok && /not a git repository/i.test(res.err)) return "This folder is not a git repository.";
  const text = res.out.trim();
  if (!text) return staged ? "(no staged changes)" : "(no unstaged changes - pass staged:true for the index, or see git_status)";
  const lines = text.split("\n");
  return lines.length > 400 ? lines.slice(0, 400).join("\n") + "\n[capped at 400 lines - pass path to narrow it]" : text;
}

export async function gitLog(root, { limit = 10 } = {}) {
  const n = Math.max(1, Math.min(50, Number(limit) || 10));
  const res = await run("git", ["log", "--oneline", "--decorate", "-n", String(n)], root, 20_000);
  if (!res.ok) {
    if (/not a git repository/i.test(res.err)) return "This folder is not a git repository.";
    if (/does not have any commits/i.test(res.err)) return "(no commits yet)";
    return "git log failed: " + (res.err.trim() || res.out.trim());
  }
  return res.out.trim() || "(no commits yet)";
}

export async function gitCommit(root, { message, all = true } = {}) {
  const msg = String(message || "").trim();
  if (!msg) throw new Error("git_commit needs a message");
  const check = await run("git", ["rev-parse", "--is-inside-work-tree"], root, 10_000);
  if (!check.ok) return "This folder is not a git repository - run `git init` first (run_command), then commit again.";
  if (all) {
    const stage = await run("git", ["add", "-A"], root, 60_000);
    if (!stage.ok) throw new Error("git add failed: " + (stage.err.trim() || stage.out.trim()));
  }
  const res = await run("git", ["commit", "-m", msg], root, 60_000);
  if (!res.ok) {
    if (/nothing to commit/i.test(res.out + res.err)) return "Nothing to commit - the working tree is clean.";
    if (/Please tell me who you are|user\.email/i.test(res.out + res.err)) {
      return "Git has no author identity. Set it once with:\n  git config --global user.name \"Your Name\"\n  git config --global user.email \"you@example.com\"";
    }
    throw new Error("git commit failed: " + (res.err.trim() || res.out.trim()));
  }
  const head = await run("git", ["log", "-1", "--oneline"], root, 15_000);
  return head.out.trim() ? "Committed: " + head.out.trim() : res.out.trim();
}

/* ------------------------------ test run ------------------------------ */

async function detectRunner(root) {
  const exists = (p) => fssync.existsSync(path.join(root, p));
  if (exists("package.json")) {
    let scripts = {};
    try {
      scripts = JSON.parse(fssync.readFileSync(path.join(root, "package.json"), "utf8")).scripts || {};
    } catch {}
    const name = ["verify", "test", "check", "lint"].find((s) => scripts[s]);
    if (name) {
      const pm = exists("pnpm-lock.yaml") ? "pnpm" : exists("yarn.lock") ? "yarn" : exists("bun.lockb") ? "bun" : "npm";
      return { command: pm, args: pm === "yarn" ? [name] : ["run", name], label: pm + " run " + name };
    }
  }
  if (exists("Cargo.toml")) return { command: "cargo", args: ["test", "--quiet"], label: "cargo test" };
  if (exists("go.mod")) return { command: "go", args: ["test", "./..."], label: "go test ./..." };
  if (exists("pyproject.toml") || exists("pytest.ini") || exists("tox.ini") || exists("tests")) {
    return { command: "python", args: ["-m", "pytest", "-q"], label: "pytest" };
  }
  if (exists("Makefile")) return { command: "make", args: ["test"], label: "make test" };
  return null;
}

export async function testRun(root, { script = "" } = {}) {
  let job = null;
  if (script) {
    const pm = fssync.existsSync(path.join(root, "pnpm-lock.yaml")) ? "pnpm" : "npm";
    job = { command: pm, args: ["run", script], label: pm + " run " + script };
  } else {
    job = await detectRunner(root);
  }
  if (!job) return "No test runner found. Add a `test` or `verify` script to package.json, or pass script:\"<npm script>\".";
  const started = Date.now();
  const res = await run(job.command, job.args, root, 600_000);
  const ms = Date.now() - started;
  const body = (res.out + (res.err ? "\n" + res.err : "")).trim();
  const verdict = res.ok ? "PASS" : "FAIL";
  const tail = body.length > 6000 ? "[first 6000 chars]\n" + body.slice(0, 6000) : body;
  return verdict + " - " + job.label + " (" + (ms / 1000).toFixed(1) + "s, exit " + res.code + ")\n\n" + (tail || "(no output)");
}

/* ------------------------------ documents ----------------------------- */

export function officeText(doc) {
  if (doc.type === "docx") {
    return doc.blocks
      .map((b) => (b.type === "table" ? b.rows.map((r) => r.join(" | ")).join("\n") : b.text))
      .join("\n");
  }
  if (doc.type === "xlsx") {
    return doc.sheets
      .map((s) => "## " + s.name + "\n" + s.rows.map((r) => r.map((c) => c.value).join("\t")).join("\n"))
      .join("\n\n");
  }
  if (doc.type === "pptx") {
    return doc.slides.map((s) => "## Slide " + s.index + "\n" + s.paragraphs.join("\n")).join("\n\n");
  }
  return "";
}

export async function officeRead(root, file) {
  const abs = path.isAbsolute(file) ? file : path.resolve(root, file);
  if (!fssync.existsSync(abs)) throw new Error("no such document: " + file);
  const kind = ext(abs);
  if (![".docx", ".xlsx", ".pptx"].includes(kind)) {
    throw new Error("office_read handles .docx, .xlsx and .pptx - for anything else use read_file");
  }
  const doc = await extractOffice(abs, kind);
  const text = officeText(doc);
  return "Read " + rel(root, abs) + " (" + doc.type + ", " + doc.size + " bytes)\n\n" + (text || "(no text found)");
}

export async function officeWrite(root, { path: p, kind = "", content = {} }) {
  if (!p) throw new Error("office_write needs a path");
  const abs = path.isAbsolute(p) ? p : path.resolve(root, p);
  const format = String(kind || ext(abs)).replace(/^\./, "").toLowerCase();
  if (!["docx", "xlsx", "pptx"].includes(format)) {
    throw new Error('office_write writes "docx", "xlsx" or "pptx" - pass kind or use a matching extension');
  }
  let data = content;
  if (typeof data === "string") {
    try {
      data = JSON.parse(data);
    } catch {
      data = format === "xlsx" ? { sheets: [{ name: "Sheet1", rows: data.split("\n").map((line) => line.split("\t")) }] } : { sections: [{ paragraphs: data.split(/\n{2,}/) }] };
    }
  }
  let buf;
  if (format === "docx") buf = writeDocx({ title: data.title || path.basename(abs, ".docx"), sections: data.sections || [] });
  else if (format === "xlsx") buf = writeXlsx({ sheets: data.sheets || [] });
  else buf = writePptx({ slides: data.slides || [] });

  await fs.mkdir(path.dirname(abs), { recursive: true });
  await fs.writeFile(abs, buf);

  const back = await extractOffice(abs, "." + format);
  const summary =
    format === "docx"
      ? back.blocks.length + " paragraph(s)"
      : format === "xlsx"
        ? back.sheets.length + " sheet(s), " + back.sheets.reduce((n, s) => n + s.rows.length, 0) + " row(s)"
        : back.slides.length + " slide(s)";
  return "Wrote " + rel(root, abs) + " (" + buf.length + " bytes, " + summary + "). Read it back with office_read to check the content.";
}

/* ------------------------------- memory ------------------------------- */

const MEMORY_LIMIT = 200;

export async function memoryOp(root, { action = "recall", key = "", value = "", query = "" } = {}) {
  const file = star(root, "memory.json");
  const store = await readJson(file, { entries: [] });
  const entries = Array.isArray(store.entries) ? store.entries : [];
  const act = String(action || "recall").toLowerCase();

  if (act === "remember" || act === "save") {
    const k = String(key || "").trim();
    const v = String(value || "").trim();
    if (!k || !v) throw new Error("memory remember needs a key and a value");
    const next = entries.filter((e) => e.key !== k);
    next.unshift({ key: k, value: v, at: new Date().toISOString() });
    await writeJson(file, { entries: next.slice(0, MEMORY_LIMIT) });
    return "Remembered: " + k + " = " + v;
  }
  if (act === "forget") {
    const k = String(key || "").trim();
    const next = entries.filter((e) => e.key !== k);
    await writeJson(file, { entries: next });
    return next.length === entries.length ? "Nothing matched " + k : "Forgot " + k;
  }
  if (act === "list") {
    if (!entries.length) return "(no memories yet - save one with action:\"remember\", key, value)";
    return entries.map((e) => "- " + e.key + ": " + e.value).join("\n");
  }
  if (act !== "recall") throw new Error('memory action must be remember, recall, list or forget');
  const needle = String(query || key || "").toLowerCase().trim();
  if (!needle) return entries.length ? entries.map((e) => "- " + e.key + ": " + e.value).join("\n") : "(no memories yet)";
  const hits = entries.filter((e) => e.key.toLowerCase().includes(needle) || e.value.toLowerCase().includes(needle));
  return hits.length ? hits.map((e) => "- " + e.key + ": " + e.value).join("\n") : "(no memory matches " + needle + ")";
}

/* ------------------------------ templates ----------------------------- */

/**
 * What the model should know up front: saved memories and the template names.
 * Synchronous because the system prompt is built once per root and cached.
 */
export function memoryPrompt(root) {
  const lines = [];
  try {
    const store = JSON.parse(fssync.readFileSync(star(root, "memory.json"), "utf8"));
    const entries = Array.isArray(store.entries) ? store.entries.slice(0, 24) : [];
    if (entries.length) {
      lines.push("Remembered about this user and project (use the memory tool to add or correct):");
      for (const e of entries) lines.push("- " + e.key + ": " + String(e.value).replace(/\s+/g, " ").slice(0, 160));
    }
  } catch {}
  try {
    const names = fssync.readdirSync(star(root, "templates")).filter((f) => f.endsWith(".md")).map((f) => f.slice(0, -3));
    if (names.length) lines.push("House style templates (read one with the template tool before writing that kind of document): " + names.slice(0, 12).join(", "));
  } catch {}
  return lines.length ? "\n" + lines.join("\n") + "\n" : "";
}

export async function templateOp(root, { action = "list", name = "", text = "" } = {}) {
  const dir = star(root, "templates");
  const act = String(action || "list").toLowerCase();
  const safe = String(name || "").trim().replace(/[^\w .-]/g, "").trim();

  if (act === "save") {
    if (!safe || !String(text || "").trim()) throw new Error("template save needs a name and the text");
    await fs.mkdir(dir, { recursive: true });
    await fs.writeFile(path.join(dir, safe + ".md"), String(text), "utf8");
    return "Saved template \"" + safe + "\" (" + text.length + " chars).";
  }
  if (act === "get") {
    if (!safe) throw new Error("template get needs a name");
    try {
      return await fs.readFile(path.join(dir, safe + ".md"), "utf8");
    } catch {
      throw new Error("no template named \"" + safe + "\" - template action:\"list\" shows what exists");
    }
  }
  if (act === "delete") {
    await fs.rm(path.join(dir, safe + ".md"), { force: true });
    return "Deleted template \"" + safe + "\".";
  }
  if (act !== "list") throw new Error("template action must be list, save, get or delete");
  let names = [];
  try {
    names = (await fs.readdir(dir)).filter((f) => f.endsWith(".md")).map((f) => f.slice(0, -3));
  } catch {}
  if (!names.length) return "(no templates yet - save one with action:\"save\", name, text to lock in a house style)";
  return names.map((n) => "- " + n).join("\n");
}

/* -------------------------------- charts ------------------------------ */

const PALETTE = ["#d97757", "#5b7f95", "#7c9463", "#b45535", "#8b877d", "#3f3c36"];

function esc(s) {
  return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function layout({ width, height, title }) {
  return { width, height, top: title ? 54 : 24, left: 48, right: 20, bottom: 44 };
}

export function chartSvg({ kind = "bar", title = "", series = [] } = {}) {
  const points = (Array.isArray(series) ? series : []).map((s) => ({ label: String(s.label ?? ""), value: Number(s.value) || 0 }));
  if (!points.length) throw new Error("chart needs series: [{label, value}, ...]");
  const W = 560;
  const H = 300;
  const box = layout({ width: W, height: H, title });
  const plotW = W - box.left - box.right;
  const plotH = H - box.top - box.bottom;
  const max = Math.max(...points.map((p) => Math.abs(p.value)), 1);
  const parts = [];
  parts.push('<svg xmlns="http://www.w3.org/2000/svg" width="' + W + '" height="' + H + '" viewBox="0 0 ' + W + " " + H + '" font-family="ui-sans-serif, system-ui, sans-serif">');
  parts.push('<rect width="' + W + '" height="' + H + '" fill="#faf9f5"/>');
  if (title) parts.push('<text x="' + box.left + '" y="32" font-size="15" font-weight="600" fill="#2b2a27">' + esc(title) + "</text>");
  parts.push('<line x1="' + box.left + '" y1="' + (H - box.bottom) + '" x2="' + (W - box.right) + '" y2="' + (H - box.bottom) + '" stroke="#e0dccf"/>');

  if (kind === "pie") {
    const total = points.reduce((n, p) => n + Math.abs(p.value), 0) || 1;
    const cx = box.left + plotW / 2;
    const cy = box.top + plotH / 2;
    const r = Math.min(plotW, plotH) / 2 - 6;
    let angle = -Math.PI / 2;
    points.forEach((p, i) => {
      const slice = (Math.abs(p.value) / total) * Math.PI * 2;
      const x1 = cx + r * Math.cos(angle);
      const y1 = cy + r * Math.sin(angle);
      const x2 = cx + r * Math.cos(angle + slice);
      const y2 = cy + r * Math.sin(angle + slice);
      const large = slice > Math.PI ? 1 : 0;
      parts.push('<path d="M' + cx + " " + cy + " L" + x1.toFixed(1) + " " + y1.toFixed(1) + " A" + r + " " + r + " 0 " + large + " 1 " + x2.toFixed(1) + " " + y2.toFixed(1) + ' Z" fill="' + PALETTE[i % PALETTE.length] + '"/>');
      angle += slice;
    });
  } else if (kind === "line") {
    const step = points.length > 1 ? plotW / (points.length - 1) : plotW;
    const d = points
      .map((p, i) => (i ? "L" : "M") + (box.left + i * step).toFixed(1) + " " + (H - box.bottom - (Math.abs(p.value) / max) * plotH).toFixed(1))
      .join(" ");
    parts.push('<path d="' + d + '" fill="none" stroke="' + PALETTE[0] + '" stroke-width="2.5" stroke-linejoin="round"/>');
    points.forEach((p, i) => {
      parts.push('<circle cx="' + (box.left + i * step).toFixed(1) + '" cy="' + (H - box.bottom - (Math.abs(p.value) / max) * plotH).toFixed(1) + '" r="3.5" fill="' + PALETTE[0] + '"/>');
    });
  } else {
    const slot = plotW / points.length;
    const barW = Math.min(64, slot * 0.6);
    points.forEach((p, i) => {
      const h = (Math.abs(p.value) / max) * plotH;
      const x = box.left + slot * i + (slot - barW) / 2;
      parts.push('<rect x="' + x.toFixed(1) + '" y="' + (H - box.bottom - h).toFixed(1) + '" width="' + barW.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="4" fill="' + PALETTE[i % PALETTE.length] + '"/>');
      parts.push('<text x="' + (box.left + slot * i + slot / 2).toFixed(1) + '" y="' + (H - box.bottom + 18) + '" font-size="11" text-anchor="middle" fill="#6f6c63">' + esc(p.label.slice(0, 12)) + "</text>");
      parts.push('<text x="' + (box.left + slot * i + slot / 2).toFixed(1) + '" y="' + (H - box.bottom - h - 7).toFixed(1) + '" font-size="11" text-anchor="middle" fill="#2b2a27">' + esc(String(p.value)) + "</text>");
    });
  }
  parts.push("</svg>");
  return parts.join("");
}

export async function chartWrite(root, args = {}) {
  const svg = chartSvg(args);
  const name = String(args.path || "").trim() || "chart-" + Date.now().toString(36) + ".svg";
  const abs = path.isAbsolute(name) ? name : star(root, "charts", name.endsWith(".svg") ? name : name + ".svg");
  await fs.mkdir(path.dirname(abs), { recursive: true });
  await fs.writeFile(abs, svg, "utf8");
  return "Saved chart to " + rel(root, abs) + " (" + svg.length + " bytes).\n\n```svg\n" + svg + "\n```";
}

/* ------------------------------ schedules ----------------------------- */
// A routine is a saved prompt plus a cadence. The model-facing tool is still
// called `schedule`; the UI calls the same store "Routines". Version 2 adds
// enabled/lastRun/history/allowTools; v1 records are upgraded on read.

const SCHEDULE_UNITS = { minute: 60_000, hour: 3_600_000, day: 86_400_000, week: 604_800_000 };
const ROUTINE_STATUS = new Set(["ok", "error", "skipped", "late", "running"]);
const ROUTINE_HISTORY_CAP = 50;
const ROUTINE_STALE_MS = 10 * 60_000;

/** The next local time at h:m strictly after `from`, skipping weekends when
 *  asked. Loops at most two weeks so a bad clock cannot hang the call. */
function nextAtTime(from, h, m, weekdaysOnly = false) {
  const d = new Date(from);
  d.setHours(h, m, 0, 0);
  for (let i = 0; i < 14; i++) {
    const weekend = d.getDay() === 0 || d.getDay() === 6;
    if (d.getTime() > from && (!weekdaysOnly || !weekend)) return d.getTime();
    d.setDate(d.getDate() + 1);
    d.setHours(h, m, 0, 0);
  }
  return from + SCHEDULE_UNITS.day;
}

export function nextRun(when, from = Date.now()) {
  const text = String(when || "").trim().toLowerCase();
  const every = /^every\s+(\d+)?\s*(minute|hour|day|week)s?$/.exec(text);
  if (every) {
    const n = Number(every[1] || 1);
    return from + n * SCHEDULE_UNITS[every[2]];
  }
  if (text === "hourly" || text === "every hour") return from + SCHEDULE_UNITS.hour;
  const daily = /^(?:every day|daily)\s+at\s+(\d{1,2}):(\d{2})$/.exec(text);
  if (daily) return nextAtTime(from, Number(daily[1]), Number(daily[2]));
  const weekdays = /^(?:every\s+)?weekdays?\s+at\s+(\d{1,2}):(\d{2})$/.exec(text);
  if (weekdays) return nextAtTime(from, Number(weekdays[1]), Number(weekdays[2]), true);
  throw new Error('routine "when" must look like "every 2 hours", "hourly", "every day at 09:00" or "weekdays at 08:30"');
}

function cryptoId() {
  return Math.random().toString(36).slice(2, 10);
}

/** v1 -> v2 without rewriting the file until the next mutation. */
function normalizeRoutine(t) {
  if (!t || !t.id) return null;
  const when = String(t.when || "").trim();
  if (!when) return null;
  let next = Number(t.next) || 0;
  if (!(next > 0)) {
    try { next = nextRun(when); } catch { next = Date.now() + SCHEDULE_UNITS.day; }
  }
  const history = Array.isArray(t.history)
    ? t.history.filter((h) => h && h.runId).map((h) => ({
        runId: String(h.runId),
        started: Number(h.started) || 0,
        ended: Number(h.ended) || 0,
        status: ROUTINE_STATUS.has(h.status) ? h.status : "ok",
        trigger: String(h.trigger || "schedule"),
        sessionId: String(h.sessionId || ""),
        error: String(h.error || ""),
      })).slice(-ROUTINE_HISTORY_CAP)
    : [];
  return {
    id: String(t.id),
    title: String(t.title || "Routine"),
    prompt: String(t.prompt || ""),
    when,
    enabled: t.enabled !== false,
    allowTools: t.allowTools === true,
    next,
    created: Number(t.created) || Date.now(),
    updated: Number(t.updated) || Number(t.created) || Date.now(),
    lastRun: Number(t.lastRun) || 0,
    lastStatus: ROUTINE_STATUS.has(t.lastStatus) ? t.lastStatus : "",
    lastError: String(t.lastError || ""),
    lastSessionId: String(t.lastSessionId || ""),
    runningSince: Number(t.runningSince) || 0,
    runNow: t.runNow === true,
    history,
  };
}

// One writer per store file: claim/finish and a user edit must never
// interleave their read-modify-write.
const routineLocks = new Map();
function withRoutineLock(file, fn) {
  const prev = routineLocks.get(file) || Promise.resolve();
  const next = prev.then(fn, fn);
  routineLocks.set(file, next.then(() => {}, () => {}));
  return next;
}

async function readRoutineStore(root) {
  const file = star(root, "schedules.json");
  const store = await readJson(file, { version: 2, tasks: [] });
  const tasks = (Array.isArray(store.tasks) ? store.tasks : []).map(normalizeRoutine).filter(Boolean);
  return { file, tasks };
}

export async function readRoutines(root) {
  return (await readRoutineStore(root)).tasks;
}

async function mutateRoutines(root, fn) {
  const file = star(root, "schedules.json");
  return withRoutineLock(file, async () => {
    const store = await readJson(file, { version: 2, tasks: [] });
    const tasks = (Array.isArray(store.tasks) ? store.tasks : []).map(normalizeRoutine).filter(Boolean);
    const out = await fn(tasks);
    await writeJson(file, { version: 2, tasks });
    return out;
  });
}

/** Create/update/delete/toggle/queue/history for one routine. Returns plain
 *  data; the tool wrapper turns it into prose. */
export async function applyRoutineOp(root, args = {}) {
  const act = String(args.action || "list").toLowerCase();
  if (act === "list") return { ok: true, tasks: await readRoutines(root) };
  return mutateRoutines(root, (tasks) => {
    const key = String(args.id || args.title || "").trim();
    const find = () => tasks.findIndex((t) => t.id === key || (!args.id && t.title === key));
    if (act === "add" || act === "create") {
      const title = String(args.title || "").trim();
      const prompt = String(args.prompt || "").trim();
      const when = String(args.when || "").trim();
      if (!title || !prompt) return { ok: false, error: "a routine needs a title and the prompt to run" };
      let next;
      try { next = nextRun(when); } catch (e) { return { ok: false, error: e.message }; }
      const routine = normalizeRoutine({
        id: cryptoId(), title, prompt, when, next, created: Date.now(),
        enabled: args.enabled !== false, allowTools: args.allowTools === true,
      });
      tasks.push(routine);
      return { ok: true, routine, tasks };
    }
    const i = find();
    if (i < 0) return { ok: false, error: "no routine matched " + (key || "(none)") };
    const t = tasks[i];
    if (act === "remove" || act === "delete") {
      tasks.splice(i, 1);
      return { ok: true, routine: t, tasks };
    }
    if (act === "enable" || act === "disable") {
      t.enabled = act === "enable";
      t.updated = Date.now();
      return { ok: true, routine: t, tasks };
    }
    if (act === "run") {
      t.runNow = true;
      t.updated = Date.now();
      return { ok: true, routine: t, queued: true, tasks };
    }
    if (act === "history") return { ok: true, routine: t, tasks };
    if (act === "update") {
      if (args.title != null && String(args.title).trim()) t.title = String(args.title).trim();
      if (args.prompt != null && String(args.prompt).trim()) t.prompt = String(args.prompt).trim();
      if (args.when != null && String(args.when).trim()) {
        try { t.when = String(args.when).trim(); t.next = nextRun(t.when); }
        catch (e) { return { ok: false, error: e.message }; }
      }
      if (args.enabled !== undefined) t.enabled = args.enabled !== false;
      if (args.allowTools !== undefined) t.allowTools = args.allowTools === true;
      t.updated = Date.now();
      return { ok: true, routine: t, tasks };
    }
    return { ok: false, error: "action must be add, update, remove, enable, disable, run, history or list" };
  });
}

/** Tasks whose slot has arrived (or that a Run now queued). */
export async function routinesDue(root, now = Date.now()) {
  const tasks = await readRoutines(root);
  return tasks.filter((t) => !t.runningSince && (t.runNow || (t.enabled && Number(t.next) > 0 && Number(t.next) <= now)));
}

/** Atomically take a routine for one run. Null when another run still holds
 *  it; a claim older than ten minutes is finalized as an error and retaken. */
export async function claimRoutine(root, id, trigger = "schedule") {
  return mutateRoutines(root, (tasks) => {
    const t = tasks.find((x) => x.id === id);
    if (!t) return null;
    const now = Date.now();
    if (t.runningSince) {
      if (now - t.runningSince < ROUTINE_STALE_MS) return null;
      const stale = (t.history || []).find((h) => h.status === "running");
      if (stale) { stale.status = "error"; stale.ended = now; stale.error = "the previous run did not finish"; }
      t.runningSince = 0;
    }
    const late = !t.runNow && t.next > 0 && now - t.next > 60_000;
    const runId = cryptoId();
    t.runNow = false;
    t.runningSince = now;
    t.lastRun = now;
    t.lastStatus = "running";
    const entry = { runId, started: now, ended: 0, status: "running", trigger, sessionId: "", error: "" };
    t.history = (t.history || []).concat(entry).slice(-ROUTINE_HISTORY_CAP);
    return { task: t, runId, late };
  });
}

/** Settle a claimed run: status, error, session and the next future slot. */
export async function finishRoutine(root, id, runId, patch = {}) {
  return mutateRoutines(root, (tasks) => {
    const t = tasks.find((x) => x.id === id);
    if (!t) return null;
    const now = Date.now();
    const status = ROUTINE_STATUS.has(patch.status) ? patch.status : "ok";
    t.runningSince = 0;
    t.lastStatus = status;
    t.lastError = String(patch.error || "");
    if (patch.sessionId) t.lastSessionId = String(patch.sessionId);
    t.updated = now;
    if (Number(patch.next) > 0) t.next = Number(patch.next);
    else { try { t.next = nextRun(t.when); } catch { /* keep the old slot */ } }
    const h = (t.history || []).find((x) => x.runId === runId) || (t.history || []).find((x) => x.status === "running");
    if (h) { h.ended = now; h.status = status; h.sessionId = String(patch.sessionId || h.sessionId || ""); h.error = String(patch.error || ""); }
    return t;
  });
}

/** Legacy name kept for one release: due routines. */
export async function schedulesDue(root, now = Date.now()) {
  return routinesDue(root, now);
}

/** Legacy name kept for one release: push one task to its next slot. */
export async function scheduleAdvance(root, id) {
  return mutateRoutines(root, (tasks) => {
    const t = tasks.find((x) => x.id === id);
    if (!t) return null;
    try { t.next = nextRun(t.when); } catch { /* keep the old slot */ }
    return t;
  });
}

/** The model-facing `schedule` tool: plain prose around applyRoutineOp. */
export async function scheduleOp(root, { action = "list", ...rest } = {}) {
  const out = await applyRoutineOp(root, { action, ...rest });
  if (!out.ok) throw new Error(out.error);
  if (action === "list") {
    if (!out.tasks || !out.tasks.length) return "(nothing scheduled)";
    return out.tasks.map((t) => "- " + t.title + " - " + t.when + (t.enabled ? "" : " (paused)") +
      " (next " + new Date(t.next).toLocaleString() + ")").join("\n");
  }
  if (out.queued) return "Queued \"" + out.routine.title + "\" to run now.";
  if (action === "history") {
    return (out.routine.history || []).map((h) => "- " + new Date(h.started).toLocaleString() + " " + h.status +
      (h.error ? " - " + h.error : "")).join("\n") || "No runs yet.";
  }
  if (action === "remove" || action === "delete") return "Removed routine " + out.routine.title + ".";
  if (action === "enable" || action === "disable") return (out.routine.enabled ? "Enabled " : "Paused ") + out.routine.title + ".";
  if (action === "add" || action === "create") {
    return "Scheduled \"" + out.routine.title + "\" for " + out.routine.when +
      " (next " + new Date(out.routine.next).toLocaleString() + ").";
  }
  return "Updated routine \"" + out.routine.title + "\" (" + out.routine.when +
    ", next " + new Date(out.routine.next).toLocaleString() + ").";
}

/* -------------------------------- trace ------------------------------- */

const TRACE_LIMIT = 4000;

export async function traceAppend(root, entry) {
  try {
    const file = star(root, "trace.jsonl");
    await fs.mkdir(path.dirname(file), { recursive: true });
    await fs.appendFile(file, JSON.stringify(entry) + "\n", "utf8");
    const stat = await fs.stat(file);
    if (stat.size > TRACE_LIMIT * 400) {
      const text = await fs.readFile(file, "utf8");
      const keep = text.split("\n").slice(-TRACE_LIMIT).join("\n");
      await fs.writeFile(file, keep, "utf8");
    }
  } catch {}
}

export async function traceTail(root, { limit = 20 } = {}) {
  const n = Math.max(1, Math.min(200, Number(limit) || 20));
  let text = "";
  try {
    text = await fs.readFile(star(root, "trace.jsonl"), "utf8");
  } catch {
    return "(no trace yet - tools record here once they run)";
  }
  const lines = text.trim().split("\n").slice(-n).reverse();
  return lines
    .map((l) => {
      try {
        const e = JSON.parse(l);
        return "- " + (e.at || "") + " " + (e.tool || "?") + " " + (e.status || "") + " " + (e.ms != null ? e.ms + "ms" : "") + (e.note ? " - " + e.note : "");
      } catch {
        return "- " + l.slice(0, 200);
      }
    })
    .join("\n");
}

/* ----------------------------- connectors ----------------------------- */

const CONNECTOR_PRESETS = {
  gmail: "email",
  "google-calendar": "calendar",
  "google-drive": "files",
  notion: "docs",
  linear: "issues",
  jira: "issues",
  slack: "chat",
  hubspot: "crm",
};

export async function connectorOp(root, { action = "list", name = "", command = "", args = [] } = {}) {
  // Connectors ARE MCP servers: write the project's real mcp.json so the
  // manager sees them. The old connectors.json was never read by the client.
  const file = star(root, "mcp.json");
  const store = await readJson(file, null);
  const base = store && typeof store === "object" ? { ...store } : {};
  let servers = base.servers && typeof base.servers === "object" ? base.servers : null;
  if (!servers) {
    const legacy = await readJson(star(root, "connectors.json"), { connectors: {} });
    const old = legacy.connectors && typeof legacy.connectors === "object" ? legacy.connectors : {};
    servers = {};
    for (const [k, v] of Object.entries(old)) {
      const cmd = [String(v.command || "").trim(), ...(Array.isArray(v.args) ? v.args.map(String) : [])].filter(Boolean);
      if (k && cmd.length) servers[k] = { command: cmd };
    }
    if (Object.keys(servers).length) await writeJson(file, { ...base, version: base.version || 1, servers });
  }
  const save = () => writeJson(file, { ...base, version: base.version || 1, servers });
  const act = String(action || "list").toLowerCase();

  if (act === "add") {
    const key = String(name || "").trim().toLowerCase();
    const cmd = [String(command || "").trim(), ...(Array.isArray(args) ? args.map(String) : String(args || "").split(/\s+/).filter(Boolean))].filter(Boolean);
    if (!key || !cmd.length) throw new Error("connector add needs a name and the MCP command to launch");
    servers[key] = { command: cmd };
    await save();
    return "Connected " + key + " (" + (CONNECTOR_PRESETS[key] || "custom") + "). The MCP manager reloads mcp.json within a second.";
  }
  if (act === "remove") {
    const key = String(name || "").trim().toLowerCase();
    delete servers[key];
    await save();
    return "Removed " + key + ".";
  }
  if (act !== "list") throw new Error("connector action must be list, add or remove");
  const configured = Object.entries(servers);
  const presets = Object.keys(CONNECTOR_PRESETS).join(", ");
  if (!configured.length) {
    return "No connectors configured. These talk to the apps your data lives in through MCP servers:\n  " + presets +
      '\nAdd one with connector action:"add", name:"gmail", command:"npx", args:["-y","@some/gmail-mcp"]';
  }
  return configured.map(([k, v]) => {
    const cmd = v.url || [].concat(v.command || []).join(" ");
    return "- " + k + " [" + (CONNECTOR_PRESETS[k] || "custom") + "] " + cmd;
  }).join("\n");
}

export { CONNECTOR_PRESETS };
