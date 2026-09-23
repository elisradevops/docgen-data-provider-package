'use strict';
import { AsyncLocalStorage } from 'async_hooks';

export interface RunContext {
  runId: string;
}

// Symbol.for uses the global symbol registry, so every duplicated copy of this file across
// the DocGen packages — hoisted or nested at any depth by npm — converges on the same
// AsyncLocalStorage instance. Keying by module identity instead would silently split into
// two stores and runId would go missing with no visible error.
const KEY = Symbol.for('elisradevops.docgen.runContext');

export const runContextStore: AsyncLocalStorage<RunContext> =
  ((globalThis as Record<symbol, unknown>)[KEY] as AsyncLocalStorage<RunContext> | undefined) ??
  ((globalThis as Record<symbol, unknown>)[KEY] = new AsyncLocalStorage<RunContext>());
