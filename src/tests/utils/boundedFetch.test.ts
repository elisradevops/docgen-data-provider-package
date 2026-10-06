import { fetchSettledBounded, fetchConcurrency } from '../../utils/boundedFetch';

describe('fetchSettledBounded', () => {
  it('fetches each unique key once and never exceeds the concurrency bound', async () => {
    let inFlight = 0;
    let maxInFlight = 0;
    const calls: string[] = [];
    const outcomes = await fetchSettledBounded(
      ['a', 'b', 'a', 'c', 'd', 'e', 'f'],
      async (key) => {
        calls.push(key);
        inFlight++;
        maxInFlight = Math.max(maxInFlight, inFlight);
        await new Promise((resolve) => setTimeout(resolve, 5));
        inFlight--;
        return key.toUpperCase();
      },
      3
    );

    expect(calls.sort()).toEqual(['a', 'b', 'c', 'd', 'e', 'f']);
    expect(maxInFlight).toBeGreaterThan(1);
    expect(maxInFlight).toBeLessThanOrEqual(3);
    expect(outcomes.get('a')).toEqual({ ok: true, value: 'A' });
  });

  it('records failures per key without rejecting or stopping the others', async () => {
    const boom = new Error('boom');
    const outcomes = await fetchSettledBounded(['a', 'b', 'c'], async (key) => {
      if (key === 'b') throw boom;
      return key;
    });

    expect(outcomes.get('b')).toEqual({ ok: false, error: boom });
    expect(outcomes.get('a')).toEqual({ ok: true, value: 'a' });
    expect(outcomes.get('c')).toEqual({ ok: true, value: 'c' });
  });

  it('resolves immediately for no keys', async () => {
    expect((await fetchSettledBounded([], async () => 1)).size).toBe(0);
  });
});

describe('fetchConcurrency', () => {
  const previous = process.env.DOCGEN_FETCH_CONCURRENCY;
  afterEach(() => {
    if (previous === undefined) delete process.env.DOCGEN_FETCH_CONCURRENCY;
    else process.env.DOCGEN_FETCH_CONCURRENCY = previous;
  });

  it('defaults to 4 and is clamped to 1..8; invalid values fall back to the default', () => {
    expect(fetchConcurrency(undefined)).toBe(4);
    expect(fetchConcurrency('')).toBe(4);
    expect(fetchConcurrency('abc')).toBe(4);
    expect(fetchConcurrency('0')).toBe(4);
    expect(fetchConcurrency('-3')).toBe(4);
    expect(fetchConcurrency('1')).toBe(1);
    expect(fetchConcurrency('6.9')).toBe(6);
    expect(fetchConcurrency('64')).toBe(8);
  });

  it('is read at call time and used when no explicit bound is given', async () => {
    process.env.DOCGEN_FETCH_CONCURRENCY = '2';
    let inFlight = 0;
    let maxInFlight = 0;

    await fetchSettledBounded(['a', 'b', 'c', 'd', 'e', 'f'], async () => {
      inFlight++;
      maxInFlight = Math.max(maxInFlight, inFlight);
      await new Promise((resolve) => setTimeout(resolve, 5));
      inFlight--;
      return 1;
    });

    expect(maxInFlight).toBe(2);
  });
});
