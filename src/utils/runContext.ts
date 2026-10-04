'use strict';
import { AsyncLocalStorage } from 'async_hooks';

export interface RunContext {
  runId: string;
  // Phase 6b — this package never sets it, only reads it: content-control (the host process
  // this package runs in-process inside) populates it via its own attachRunContext, and this
  // package's DiagnosticsTransport reads the same in-process ALS store.
  captureMode?: 'verbose' | 'retain-on-failure';
  // Phase 7b — same read-only treatment as captureMode above.
  docType?: string;
  // Phase 7c — same read-only treatment; set by content-control's attachRunContext via
  // x-docgen-project, visible here through the shared in-process ALS store.
  project?: string;
}

// Symbol.for uses the global symbol registry, so every duplicated copy of this file across
// the DocGen packages — hoisted or nested at any depth by npm — converges on the same
// AsyncLocalStorage instance. Keying by module identity instead would silently split into
// two stores and runId would go missing with no visible error.
const KEY = Symbol.for('elisradevops.docgen.runContext');

export const runContextStore: AsyncLocalStorage<RunContext> =
  ((globalThis as Record<symbol, unknown>)[KEY] as AsyncLocalStorage<RunContext> | undefined) ??
  ((globalThis as Record<symbol, unknown>)[KEY] = new AsyncLocalStorage<RunContext>());
