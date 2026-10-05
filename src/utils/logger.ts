'use strict';
import * as winston from 'winston';
import * as fs from 'fs';
import * as path from 'path';
import * as Transport from 'winston-transport';
import { runContextStore } from './runContext';
import { getLogSink, DiagnosticEvent, CONTEXT_LIMITS } from './logSink';
let logger: winston.Logger;

// Merges the ambient runId (set by content-control's request middleware, once Phase 3 lands
// there) into every log record. A never-populated store is a harmless no-op — this package
// only reads it, never sets it.
export const withRunContext = winston.format((info) => {
  const store = runContextStore.getStore();
  if (store?.runId) (info as Record<string, unknown>).runId = store.runId;
  // Phase 7b — same ambient, per-run treatment as runId.
  if (store?.docType) (info as Record<string, unknown>).docType = store.docType;
  // Phase 7c — same read-only treatment; populated by content-control's attachRunContext.
  if (store?.project) (info as Record<string, unknown>).project = store.project;
  // Which generation stage / content control is being served — the transport copies these onto
  // the event. An explicit value passed in the log call's own metadata wins.
  const target = info as Record<string, unknown>;
  if (store?.step && target.step === undefined) target.step = store.step;
  if (store?.contentControlType && target.contentControlType === undefined) target.contentControlType = store.contentControlType;
  if (store?.contentControlTitle && target.contentControlTitle === undefined) target.contentControlTitle = store.contentControlTitle;
  return info;
});

// Defense-in-depth: scrubs known-sensitive keys out of any object attached to a log call
// (meta, splat, an Error's own enumerable props — e.g. AxiosError.toJSON()'s
// config.auth.password / config.headers.Authorization, which is exactly how a PAT was
// leaking via `logger.error(JSON.stringify(error))` on ADO REST failures). Does not
// redact secrets interpolated into a message string — that's the corresponding call-site
// fixes' job; this is a backstop for structured fields.
// "accesskey" (not just "minioaccesskey") so this also catches an *AccessKeyId-style field —
// a real gap found in api-gate's copy during the Phase 5 manifest work: only the *SecretKey
// sibling was covered before, via "secret". Applied here too to keep the four copies in sync.
// Word-based, not substring: the key is split into lower-case words (camelCase, kebab-case and
// snake_case alike) and judged by how it *ends*. A substring test matched "pat" inside
// "areaPath" / "path" / "patch" and blanked ordinary ADO fields; this keeps `accessToken`,
// `x-docgen-ingest-token`, `minioSecretKey`, `PAT`, `password` redacted while `areaPath` stays.
const SENSITIVE_TAILS = ['token', 'password', 'secret', 'authorization', 'accesskey', 'secretkey', 'apikey'];
const SENSITIVE_LAST_WORDS = new Set(['pat', 'pwd', 'cookie', 'tokens', 'passwords', 'secrets']);
export function isSensitiveKey(key: string): boolean {
  const words = key
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter(Boolean);
  if (words.length === 0) return false;
  if (SENSITIVE_LAST_WORDS.has(words[words.length - 1])) return true;
  const tail = words.slice(-2).join('');
  return SENSITIVE_TAILS.some((t) => tail.endsWith(t));
}

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
export function redactValue(value: unknown, depth = 0, seen = new WeakSet<object>()): unknown {
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
    out[key] = isSensitiveKey(key) ? '[REDACTED]' : redactValue(read.value, depth + 1, seen);
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
        isSensitiveKey(key)
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

// A template literal's implicit ToString throws on a Symbol value (unlike String(), which
// calls Symbol.prototype.toString() explicitly) — logger.error(Symbol('x')) would otherwise
// crash inside this formatter itself, the one place in the pipeline with no try/catch around
// it. Symbol.toString() itself never throws, so this needs no further guarding.
const safeMessageString = (value: unknown): string =>
  typeof value === 'symbol' ? value.toString() : String(value);

// A failed ADO request's description (helpers/requestContext.ts), attached to the thrown error
// and merged onto the record by winston. In text mode — the default, which prints only the
// message — this is the only way the URL reaches stdout for a record whose own message doesn't
// carry it. Never throws: this runs inside the formatter, which has no try/catch around it.
const requestSuffix = (ctx: unknown): string => {
  try {
    if (!ctx || typeof ctx !== 'object') return '';
    const c = ctx as Record<string, unknown>;
    if (typeof c.url !== 'string') return '';
    const method = typeof c.method === 'string' ? `${c.method} ` : '';
    const status = typeof c.status === 'number' ? ` -> ${c.status}` : '';
    return ` [${method}${c.url}${status}]`;
  } catch {
    return '';
  }
};

const textFormat = winston.format.printf(
  (info) => `${info.timestamp} - ${info.level}: ${safeMessageString(info.message)}${requestSuffix(info.adoRequest)}`
);

// Bounded so one oversized message/stack can't produce an unbounded LogEvent document.
const MAX_MESSAGE_LEN = 2000;
const MAX_STACK_LEN = 4000;
function clamp(value: unknown, max: number): string | undefined {
  if (typeof value !== 'string') return undefined;
  return value.length > max ? value.slice(0, max) : value;
}

// Re-validates the request description at the transport instead of trusting its shape — it
// arrives via an Error's own properties, which anything upstream could have set. Allowlisted
// keys only, each type-checked and bounded.
function pickContext(raw: unknown): DiagnosticEvent['context'] {
  if (!raw || typeof raw !== 'object') return undefined;
  const c = raw as Record<string, unknown>;
  const context: NonNullable<DiagnosticEvent['context']> = {};
  const method = clamp(c.method, CONTEXT_LIMITS.method);
  const url = clamp(c.url, CONTEXT_LIMITS.url);
  const requestBody = clamp(c.requestBody, CONTEXT_LIMITS.requestBody);
  const responseExcerpt = clamp(c.responseExcerpt, CONTEXT_LIMITS.responseExcerpt);
  if (method) context.method = method;
  if (url) context.url = url;
  if (requestBody) context.requestBody = requestBody;
  if (responseExcerpt) context.responseExcerpt = responseExcerpt;
  if (typeof c.status === 'number' && Number.isFinite(c.status)) context.status = c.status;
  if (typeof c.attempt === 'number' && Number.isFinite(c.attempt)) context.attempt = c.attempt;
  return Object.keys(context).length ? context : undefined;
}

// Ships warn/error records to whatever LogSink the host process (docgen-content-control's
// index.ts) has installed — this package never installs one itself, only reads it. This is
// the dominant source of real ADO error records (tfs.ts's executeWithRetry is the single
// static axios instance handling all Azure DevOps traffic), so without this transport most
// real generation failures either never reach the store or get misattributed to
// dg-content-control. Reads info *after* the rest of the format chain has already run —
// redact() has already scrubbed it by the time this transport sees it. Never throws, never
// blocks the call path, and never logs through `logger` itself (that would recurse back
// into this same transport) — failures go to console.error only.
const DIAGNOSTICS_CAPTURE_ENABLED = (process.env.DIAGNOSTICS_CAPTURE_ENABLED || 'true').toLowerCase() !== 'false';
// Exported for tests, the same treatment as redact()/withRunContext() — so a test pipeline
// can exercise this transport without going through the real singleton logger's env-gated
// json()/textFormat branch.
export class DiagnosticsTransport extends Transport {
  log(info: Record<string, unknown>, callback: () => void): void {
    setImmediate(() => this.emit('logged', info));
    try {
      const level = info.level;
      const isWarnOrError = level === 'warn' || level === 'error';
      // Phase 6b — the logger's own level gate is now 'debug' (see createLogger below), so a
      // debug/info call reaches this transport regardless of mode; this is the one place that
      // decides whether it's actually persisted, keyed off the per-run capture mode rather
      // than a process-wide setting — concurrent runs in 'normal' mode are unaffected by a
      // sibling run's 'verbose'/'retain-on-failure' opt-in.
      const captureMode = runContextStore.getStore()?.captureMode;
      const isCapturedDebugOrInfo =
        (level === 'debug' || level === 'info') && (captureMode === 'verbose' || captureMode === 'retain-on-failure');
      if (DIAGNOSTICS_CAPTURE_ENABLED && (isWarnOrError || isCapturedDebugOrInfo)) {
        // winston.errors({stack:true}) merges an Error's own enumerable properties (stack,
        // and anything else the call site set, e.g. `err.code`) directly onto `info` — there
        // is no separate nested info.err. `info.stack`'s presence is the only reliable signal
        // that this record came from `logger.error('...', err)`/`logger.error(err)` rather
        // than a plain string message. When present, info.message is already the
        // stable-message-plus-error-message string winston produced, so err.message reuses it.
        const hasErr = typeof info.stack === 'string';
        const message = clamp(info.message, MAX_MESSAGE_LEN) ?? '';
        const event: DiagnosticEvent = {
          ts: typeof info.timestamp === 'string' ? info.timestamp : new Date().toISOString(),
          level: String(info.level),
          service: String(info.service ?? '@elisra-devops/docgen-data-provider'),
          version: String(info.version ?? 'unknown'),
          runId: typeof info.runId === 'string' ? info.runId : undefined,
          docType: typeof info.docType === 'string' ? info.docType : undefined,
          step: typeof info.step === 'string' ? info.step : undefined,
          contentControlType: typeof info.contentControlType === 'string' ? info.contentControlType : undefined,
          contentControlTitle: typeof info.contentControlTitle === 'string' ? info.contentControlTitle : undefined,
          project: typeof info.project === 'string' ? info.project : undefined,
          userId: typeof info.userId === 'string' ? info.userId : undefined,
          message,
          err: hasErr
            ? {
                // The error's own message when the call site supplied it (describeError); else
                // the combined winston message, which is all a bare logger.error(err) leaves.
                message: clamp(info.errMessage, MAX_MESSAGE_LEN) ?? message,
                code: typeof info.code === 'string' ? info.code : undefined,
                stack: clamp(info.stack, MAX_STACK_LEN),
              }
            : undefined,
          context: pickContext(info.adoRequest),
          retainPending: !isWarnOrError && captureMode === 'retain-on-failure' ? true : undefined,
        };
        getLogSink()?.push(event);
      }
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error('DiagnosticsTransport failed to push event', e);
    }
    callback();
  }
}

// LOG_FORMAT=json is the eventual default (structured stdout a collector can parse), but
// stays opt-in for one release so switching is a config change, not an image rebuild, if
// anyone reading raw console output needs to roll back.
const useJson = (process.env.LOG_FORMAT || 'text').toLowerCase() === 'json';

// Phase 6b — the Logger's own level gate must be 'debug' (winston's most permissive) so a
// debug/info call always reaches every transport's log(); a static per-logger level can't
// depend on the ALS store's per-run capture mode, so DiagnosticsTransport has to be the one
// deciding what's actually persisted. The Console transport below gets its own explicit
// level so stdout's default behavior (info/warn/error unless LOG_LEVEL=debug) is unchanged —
// this only spends the format-chain cost on debug() calls that previously short-circuited
// for free; info/warn/error volume (Phase 4's focus) is unaffected either way.
const CONSOLE_LOG_LEVEL = process.env.LOG_LEVEL || 'info';

// Drops a record before the (comparatively expensive) redact/splat/format work when nothing
// would consume it: stdout wouldn't print it, and DiagnosticsTransport wouldn't persist it
// (only debug/info under a verbose/retain-on-failure run are captured below warn). The logger's
// own level gate must stay 'debug' for that capture, so this is the cheap early exit instead.
const LEVEL_RANK: Record<string, number> = { error: 0, warn: 1, info: 2, http: 3, verbose: 4, debug: 5, silly: 6 };
export const skipUncaptured = winston.format((info) => {
  const printRank = LEVEL_RANK[CONSOLE_LOG_LEVEL];
  const rank = LEVEL_RANK[String(info.level)];
  if (printRank === undefined || rank === undefined || rank <= printRank) return info;
  const mode = runContextStore.getStore()?.captureMode;
  const captured =
    DIAGNOSTICS_CAPTURE_ENABLED &&
    (info.level === 'debug' || info.level === 'info') &&
    (mode === 'verbose' || mode === 'retain-on-failure');
  return captured ? info : false;
});

logger = winston.createLogger({
  level: 'debug',
  defaultMeta: { service: '@elisra-devops/docgen-data-provider', version: readOwnVersion() },
  format: useJson
    ? winston.format.combine(
        skipUncaptured(),
        winston.format.errors({ stack: true }),
        winston.format.timestamp(),
        withRunContext(),
        redact(),
        winston.format.splat(),
        winston.format.json()
      )
    : winston.format.combine(
        skipUncaptured(),
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
  transports: [new winston.transports.Console({ level: CONSOLE_LOG_LEVEL }), new DiagnosticsTransport()],
});

export default logger;
