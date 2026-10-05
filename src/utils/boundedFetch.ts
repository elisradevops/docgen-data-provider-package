export type Settled<T> = { ok: true; value: T } | { ok: false; error: unknown };

/**
 * Runs `fetcher` for each unique key with a small bounded pool and returns every outcome by key.
 * Callers that then walk their data in the original order read from the map, so ordering and
 * per-item error handling stay exactly as in a sequential loop while the round trips overlap.
 * It never rejects: a failure is recorded as `{ ok: false, error }` for the caller to rethrow or log.
 */
export async function fetchSettledBounded<T>(
  keys: string[],
  fetcher: (key: string) => Promise<T>,
  concurrency = 4
): Promise<Map<string, Settled<T>>> {
  const outcomes = new Map<string, Settled<T>>();
  const pending = Array.from(new Set(keys));
  let next = 0;
  const worker = async () => {
    while (next < pending.length) {
      const key = pending[next++];
      try {
        outcomes.set(key, { ok: true, value: await fetcher(key) });
      } catch (error) {
        outcomes.set(key, { ok: false, error });
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, pending.length) }, worker));
  return outcomes;
}
