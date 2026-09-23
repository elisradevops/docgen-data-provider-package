'use strict';
import * as winston from 'winston';
import * as fs from 'fs';
import * as path from 'path';
import { runContextStore } from './runContext';
let logger: winston.Logger;

// Merges the ambient runId (set by content-control's request middleware, once Phase 3 lands
// there) into every log record. A never-populated store is a harmless no-op — this package
// only reads it, never sets it.
export const withRunContext = winston.format((info) => {
  const runId = runContextStore.getStore()?.runId;
  if (runId) (info as Record<string, unknown>).runId = runId;
  return info;
});

// Defense-in-depth: scrubs known-sensitive keys out of any object attached to a log call
// (meta, splat, an Error's own enumerable props — e.g. AxiosError.toJSON()'s
// config.auth.password / config.headers.Authorization, which is exactly how a PAT was
// leaking via `logger.error(JSON.stringify(error))` on ADO REST failures). Does not
// redact secrets interpolated into a message string — that's the corresponding call-site
// fixes' job; this is a backstop for structured fields.
const SENSITIVE_KEY = /token|pat|password|secret|authorization|minioaccesskey|miniosecretkey/i;

// A per-key try/catch on the *read*, not just around the whole loop: a getter that throws
// (e.g. `{ get boom() { throw ... } }`) would otherwise abort redaction for every remaining
// key in the same object, or — worse — survive redaction untouched and re-throw later, inside
// winston's json() formatter, which is not wrapped in anything. redactValue always returns a
// freshly built plain object/array, never the original reference, so once a value has passed
// through here nothing downstream can re-trigger a throwing accessor on it.
function safeRead(obj: Record<string, unknown>, key: string): { ok: true; value: unknown } | { ok: false } {
  try {
    return { ok: true, value: obj[key] };
  } catch {
    return { ok: false };
  }
}
function redactValue(value: unknown, depth = 0, seen = new WeakSet<object>()): unknown {
  if (depth > 6 || value === null || typeof value !== 'object') return value;
  if (seen.has(value as object)) return '[Circular]';
  seen.add(value as object);
  if (Array.isArray(value)) return value.map((v) => redactValue(v, depth + 1, seen));
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>)) {
    const read = safeRead(value as Record<string, unknown>, key);
    if (!read.ok) {
      out[key] = '[Unreadable property]';
      continue;
    }
    out[key] = SENSITIVE_KEY.test(key) ? '[REDACTED]' : redactValue(read.value, depth + 1, seen);
  }
  return out;
}
export const redact = winston.format((info) => {
  const target = info as Record<string, unknown>;
  for (const key of Object.keys(target)) {
    // level/timestamp are always plain strings winston sets itself — never touched. message is
    // NOT exempt: when logger.error(x) is called with a single non-string/non-Error argument,
    // winston nests the whole thing under info.message (this is exactly the shape hit by a raw
    // object argument), so message can carry a throwing getter or a sensitive key just as much
    // as any other field — redactValue already no-ops on a plain string, so including it here
    // costs nothing for the common case and closes that gap for the uncommon one.
    if (key === 'level' || key === 'timestamp') continue;
    const read = safeRead(target, key);
    const value = !read.ok
      ? '[Unreadable property]'
      : // A top-level primitive (e.g. token: "abc") has no children for redactValue to walk into —
        // the key itself has to be checked here too, not just inside the recursive object walk.
        SENSITIVE_KEY.test(key)
        ? '[REDACTED]'
        : redactValue(read.value);
    try {
      target[key] = value;
    } catch {
      // The original property could be a getter-only accessor with no setter, which throws on
      // plain assignment in strict mode — redefine it outright rather than skip it.
      Object.defineProperty(target, key, { value, writable: true, enumerable: true, configurable: true });
    }
  }
  return info;
});

// Resolves the package's own version for defaultMeta. Node-only (this package is currently
// only ever consumed by docgen-content-control at runtime — the browser transport below is
// unused/commented out), so fs/path are safe here.
function readOwnVersion(): string {
  const candidates = [
    path.resolve(__dirname, '../package.json'),
    path.resolve(__dirname, '../../package.json'),
    path.resolve(process.cwd(), 'package.json'),
  ];
  for (const candidate of candidates) {
    try {
      if (!fs.existsSync(candidate)) continue;
      return JSON.parse(fs.readFileSync(candidate, 'utf8')).version || 'unknown';
    } catch {
      // try next candidate
    }
  }
  return 'unknown';
}

const textFormat = winston.format.printf((info) => `${info.timestamp} - ${info.level}: ${info.message}`);

// LOG_FORMAT=json is the eventual default (structured stdout a collector can parse), but
// stays opt-in for one release so switching is a config change, not an image rebuild, if
// anyone reading raw console output needs to roll back.
const useJson = (process.env.LOG_FORMAT || 'text').toLowerCase() === 'json';

logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  defaultMeta: { service: '@elisra-devops/docgen-data-provider', version: readOwnVersion() },
  format: useJson
    ? winston.format.combine(
        winston.format.errors({ stack: true }),
        winston.format.timestamp(),
        withRunContext(),
        redact(),
        winston.format.splat(),
        winston.format.json()
      )
    : winston.format.combine(
        winston.format.errors({ stack: true }),
        winston.format.timestamp(),
        withRunContext(),
        redact(),
        winston.format.splat(),
        textFormat
      ),
  // Stdout only — no File transports, consistent with the other DocGen services (no mounted
  // log volume anywhere in this deployment, so a File transport is invisible to `kubectl logs`
  // and can throw at import time under a read-only root filesystem).
  transports: [new winston.transports.Console()],
});

export default logger;
