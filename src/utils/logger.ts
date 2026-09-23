'use strict';
import * as winston from 'winston';
import BrowserConsole from 'winston-transport-browserconsole';
let logger: winston.Logger;

const logFormat = winston.format.printf((info) => `${info.timestamp} - ${info.level}: ${info.message}`);

// Defense-in-depth: scrubs known-sensitive keys out of any object attached to a log call
// (meta, splat, an Error's own enumerable props — e.g. AxiosError.toJSON()'s
// config.auth.password / config.headers.Authorization, which is exactly how a PAT was
// leaking via `logger.error(JSON.stringify(error))` on ADO REST failures). Does not
// redact secrets interpolated into a message string — that's the corresponding call-site
// fixes' job; this is a backstop for structured fields.
const SENSITIVE_KEY = /token|pat|password|secret|authorization|minioaccesskey|miniosecretkey/i;
function redactValue(value: unknown, depth = 0, seen = new WeakSet<object>()): unknown {
  if (depth > 6 || value === null || typeof value !== 'object') return value;
  if (seen.has(value as object)) return '[Circular]';
  seen.add(value as object);
  if (Array.isArray(value)) return value.map((v) => redactValue(v, depth + 1, seen));
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>)) {
    out[key] = SENSITIVE_KEY.test(key)
      ? '[REDACTED]'
      : redactValue((value as Record<string, unknown>)[key], depth + 1, seen);
  }
  return out;
}
export const redact = winston.format((info) => {
  try {
    for (const key of Object.keys(info)) {
      if (key === 'level' || key === 'message' || key === 'timestamp') continue;
      // A top-level primitive (e.g. token: "abc") has no children for redactValue to walk into —
      // the key itself has to be checked here too, not just inside the recursive object walk.
      (info as Record<string, unknown>)[key] = SENSITIVE_KEY.test(key)
        ? '[REDACTED]'
        : redactValue((info as Record<string, unknown>)[key]);
    }
  } catch {
    // Redaction must never break logging itself.
  }
  return info;
});
// if (typeof window === "undefined") {
//   let logsPath = process.env.logs_path || "./logs/";

//   logger = winston.createLogger({
//     format: winston.format.timestamp(),
//     level: "silly",
//     transports: [
//       new winston.transports.File({
//         filename: `${logsPath}azure-rest-api-errors.log`,
//         level: "error",
//         format: logFormat,
//       }),
//       new winston.transports.File({
//         filename: `${logsPath}azure-rest-api-all.log`,
//         format: logFormat,
//       }),
//       new winston.transports.Console({ format: logFormat, level: "debug" }),
//     ],
//   });
// } else {
logger = winston.createLogger({
  format: winston.format.combine(winston.format.timestamp(), redact()),
  level: 'silly',
  transports: [
    new winston.transports.Console({ format: logFormat, level: 'debug' }),
    // new BrowserConsole({ format: logFormat, level: 'debug' }),
  ],
});
// }

export default logger;
