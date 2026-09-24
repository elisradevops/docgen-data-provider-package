'use strict';

// The one queryable-error-store event shape, emitted by DiagnosticsTransport (logger.ts).
// This package never installs a sink itself — it only reads whatever the host process
// (docgen-content-control's index.ts) installed. err is intentionally narrow —
// {message, code, stack} — per the settled Phase 6 schema decision: no generic
// extra-fields bucket.
export interface DiagnosticEvent {
  ts: string;
  level: string;
  service: string;
  version: string;
  runId?: string;
  step?: string;
  contentControlType?: string;
  contentControlTitle?: string;
  project?: string;
  userId?: string;
  message: string;
  err?: { message: string; code?: string; stack?: string };
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
