import { redactValue } from '../utils/logger';
import { CONTEXT_LIMITS } from '../utils/logSink';

// What a failed Azure DevOps request looks like to whoever is triaging it. Built once, at the
// TFSServices choke point, and attached to the thrown error — so every caller that logs that
// error (however it words its own message) carries the same sanitized request description
// into the log record and the diagnostics store, instead of only the call sites that happened
// to print the URL themselves.
export interface AdoRequestContext {
  method?: string;
  url?: string;
  status?: number;
  attempt?: number;
  requestBody?: string;
  responseExcerpt?: string;
}

const MAX_URL_LEN = CONTEXT_LIMITS.url;
const MAX_BODY_LEN = CONTEXT_LIMITS.requestBody;
const MAX_STRING_VALUE_LEN = 500;
const MAX_ARRAY_ITEMS = 50;
const MAX_PATCH_OPS = 20;
const MAX_RESPONSE_LEN = CONTEXT_LIMITS.responseExcerpt;
const MAX_TEXT_RESPONSE_LEN = 200;

// Exact names, not a substring match — "key" must not mask a legitimate "keywords" parameter.
const SENSITIVE_PARAM = /^(access_?token|token|pat|password|pwd|secret|sig|signature|api_?key|key|authorization|auth)$/i;

const truncate = (text: string, max: number): string => (text.length > max ? `${text.slice(0, max)}…` : text);

/** Strips credentials embedded in the URL and masks sensitive query parameter values. */
export function sanitizeUrl(raw: unknown): string | undefined {
  if (typeof raw !== 'string' || !raw) return undefined;
  try {
    const parsed = new URL(raw);
    parsed.username = '';
    parsed.password = '';
    // Rewrite only the sensitive values in the raw query string — re-serializing through
    // URLSearchParams would re-encode every other parameter ($depth, %20, +).
    if (parsed.search) {
      parsed.search = parsed.search
        .slice(1)
        .split('&')
        .map((pair) => {
          const eq = pair.indexOf('=');
          if (eq < 0) return pair;
          let name = pair.slice(0, eq);
          try {
            name = decodeURIComponent(name);
          } catch {
            // keep the raw name
          }
          return SENSITIVE_PARAM.test(name) ? `${pair.slice(0, eq)}=[REDACTED]` : pair;
        })
        .join('&');
    }
    return truncate(parsed.toString(), MAX_URL_LEN);
  } catch {
    return truncate(raw, MAX_URL_LEN);
  }
}

/**
 * Replaces each single-quoted WIQL literal with its length — the query structure (fields,
 * operators, numbers, @macros) stays readable, free-text values never reach the logs.
 * `''` is WIQL's escaped quote.
 */
export function maskWiqlLiterals(text: string): string {
  return text.replace(/'(?:[^']|'')*'/g, (m) => `'[${m.length - 2} chars]'`);
}

function isBinaryLike(data: unknown): boolean {
  return (
    (typeof Buffer !== 'undefined' && Buffer.isBuffer(data)) ||
    data instanceof ArrayBuffer ||
    ArrayBuffer.isView(data as ArrayBufferView)
  );
}

function isJsonPatch(data: unknown): data is Array<{ op: string; path: string }> {
  return (
    Array.isArray(data) &&
    data.length > 0 &&
    data.every(
      (item) =>
        item &&
        typeof item === 'object' &&
        typeof (item as Record<string, unknown>).op === 'string' &&
        typeof (item as Record<string, unknown>).path === 'string'
    )
  );
}

function clampValues(value: unknown, depth = 0): unknown {
  if (typeof value === 'string') return truncate(value, MAX_STRING_VALUE_LEN);
  if (depth > 6 || value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) {
    const items = value.slice(0, MAX_ARRAY_ITEMS).map((v) => clampValues(v, depth + 1));
    return value.length > MAX_ARRAY_ITEMS ? [...items, `…(+${value.length - MAX_ARRAY_ITEMS} more)`] : items;
  }
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>)) {
    out[key] = clampValues((value as Record<string, unknown>)[key], depth + 1);
  }
  return out;
}

/**
 * A short, safe description of a request body — enough to see *what was asked* (the WIQL
 * text, the id list) without ever carrying credentials or work-item content:
 *  - GET/HEAD never have one.
 *  - JSON-Patch arrays (work item / test case updates) keep only `op` + `path`; the `value`s
 *    are real field content and are dropped.
 *  - a WIQL `query` keeps its structure but its quoted literals are masked.
 *  - other JSON goes through redactValue (credential-looking keys scrubbed at any depth),
 *    with every string and the total length bounded.
 *  - anything that isn't plain JSON is described by type and size only.
 */
export function summarizeBody(data: unknown, method?: string): string | undefined {
  const verb = String(method || '').toUpperCase();
  if (verb === 'GET' || verb === 'HEAD') return undefined;
  if (data === undefined || data === null || data === '') return undefined;

  if (typeof data === 'string') {
    try {
      const parsed = JSON.parse(data);
      return typeof parsed === 'string' ? `[text, ${data.length} chars]` : summarizeBody(parsed, method);
    } catch {
      return `[text, ${data.length} chars]`;
    }
  }
  if (typeof data !== 'object') return `[${typeof data}]`;
  if (isBinaryLike(data)) {
    const size = (data as { byteLength?: number }).byteLength;
    return `[binary${typeof size === 'number' ? `, ${size} bytes` : ''}]`;
  }
  if (typeof (data as { pipe?: unknown }).pipe === 'function') return '[stream]';
  if (typeof FormData !== 'undefined' && data instanceof FormData) return '[form-data]';

  if (isJsonPatch(data)) {
    const ops = data.slice(0, MAX_PATCH_OPS).map((o) => `${o.op} ${o.path}`);
    const more = data.length > MAX_PATCH_OPS ? ` …(+${data.length - MAX_PATCH_OPS} more)` : '';
    return truncate(`json-patch (${data.length} ops, values omitted): ${ops.join('; ')}${more}`, MAX_BODY_LEN);
  }
  if (Array.isArray(data) ? data.length === 0 : Object.keys(data).length === 0) return undefined;

  try {
    const body =
      !Array.isArray(data) && typeof (data as { query?: unknown }).query === 'string'
        ? { ...(data as Record<string, unknown>), query: maskWiqlLiterals((data as { query: string }).query) }
        : data;
    return truncate(JSON.stringify(clampValues(redactValue(body))), MAX_BODY_LEN);
  } catch {
    return '[unserializable body]';
  }
}

/** The server's own explanation of the failure (ADO puts it in `message`), bounded. */
export function summarizeResponse(data: unknown): string | undefined {
  if (data === undefined || data === null) return undefined;
  if (typeof data === 'string') return truncate(data, MAX_TEXT_RESPONSE_LEN);
  if (isBinaryLike(data)) return undefined;
  if (typeof (data as { pipe?: unknown }).pipe === 'function') return '[stream]';
  try {
    const message = (data as { message?: unknown }).message;
    if (typeof message === 'string') return truncate(message, MAX_RESPONSE_LEN);
    return truncate(JSON.stringify(redactValue(data)), MAX_RESPONSE_LEN);
  } catch {
    return '[unserializable response data]';
  }
}

export function buildAdoRequestContext(
  error: any,
  request: { method?: string; url?: string; data?: unknown },
  attempt?: number
): AdoRequestContext {
  const method = String(request.method || error?.config?.method || 'get').toUpperCase();
  const context: AdoRequestContext = { method, url: sanitizeUrl(request.url ?? error?.config?.url) };
  const status = Number(error?.response?.status);
  if (Number.isFinite(status) && status > 0) context.status = status;
  if (attempt !== undefined) context.attempt = attempt;
  const requestBody = summarizeBody(request.data, method);
  if (requestBody) context.requestBody = requestBody;
  const responseExcerpt = summarizeResponse(error?.response?.data);
  if (responseExcerpt) context.responseExcerpt = responseExcerpt;
  return context;
}

/**
 * Attaches the context as an enumerable own property: winston's errors() / log() merge an
 * Error's own enumerable properties onto the log record, which is how the transport and the
 * text formatter find it without every call site having to pass it along.
 */
export function annotateAdoError<T>(error: T, context: AdoRequestContext): T {
  if (error && typeof error === 'object') {
    try {
      Object.defineProperty(error, 'adoRequest', { value: context, enumerable: true, configurable: true, writable: true });
    } catch {
      // A frozen/exotic error object just goes unannotated — never worth failing the request path.
    }
  }
  return error;
}

/**
 * The log metadata for a caught error: message, stack, code, and the request description when
 * the error came through TFSServices. Deliberately NOT the error object itself — passing a
 * whole AxiosError would merge its entire config/request/response graph onto the record.
 * winston concatenates meta.message onto the call's own message, so
 * `logger.error('Could not fetch X:', describeError(err))` reads the same as the old
 * `logger.error(\`Could not fetch X: ${err.message}\`)` — plus the request, which a message
 * string can't carry.
 */
export function describeError(error: unknown): Record<string, unknown> {
  if (!error || typeof error !== 'object') return { message: String(error) };
  const e = error as Record<string, unknown>;
  const meta: Record<string, unknown> = { message: e.message, stack: e.stack };
  if (typeof e.code === 'string') meta.code = e.code;
  if (e.adoRequest) meta.adoRequest = e.adoRequest;
  return meta;
}
