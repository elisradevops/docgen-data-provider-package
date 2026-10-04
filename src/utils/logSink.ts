'use strict';

// The one queryable-error-store event shape, emitted by DiagnosticsTransport (logger.ts).
// This package never installs a sink itself — it only reads whatever the host process
// (docgen-content-control's index.ts) installed. err stays narrow — {message, code, stack}.
// `context` is the one deliberate exception to "no extra-fields bucket": a sanitized
// description of the failed outbound request (method, url, status, attempt, request-body
// summary, response excerpt) — without it a 404 in the dashboard can't be traced to the call
// that produced it. A fixed, allowlisted shape, not a free-form bag.
// Bounds for DiagnosticEvent.context. Single source for this package (requestContext.ts and the
// logger transport); docgen-api-gate's DiagnosticsController.sanitizeContext re-validates at
// ingest with its own copy — keep the two in step.
export const CONTEXT_LIMITS = { method: 10, url: 1000, requestBody: 2000, responseExcerpt: 300 } as const;

export interface DiagnosticEvent {
  ts: string;
  level: string;
  service: string;
  version: string;
  runId?: string;
  docType?: string;
  step?: string;
  contentControlType?: string;
  contentControlTitle?: string;
  project?: string;
  userId?: string;
  message: string;
  err?: { message: string; code?: string; stack?: string };
  context?: {
    method?: string;
    url?: string;
    status?: number;
    attempt?: number;
    requestBody?: string;
    responseExcerpt?: string;
  };
  // Phase 6b — set on a debug/info event captured under retain-on-failure. Deleted by
  // api-gate at the run's one success point; left alone (and thus permanent, subject to the
  // normal TTL) if the run fails.
  retainPending?: boolean;
}

export interface LogSink {
  push(event: DiagnosticEvent): void;
}

// Symbol.for so this package's copy converges on the same installed sink as
// docgen-content-control's (the host process this package runs inside) and
// docgen-dg-skins-package's, the same reasoning as runContext.ts's AsyncLocalStorage. A
// never-populated sink is a harmless no-op — this package only reads it, never installs one.
const KEY = Symbol.for('elisradevops.docgen.logSink');

export function getLogSink(): LogSink | undefined {
  return (globalThis as Record<symbol, unknown>)[KEY] as LogSink | undefined;
}

export function installLogSink(sink: LogSink): void {
  (globalThis as Record<symbol, unknown>)[KEY] = sink;
}
