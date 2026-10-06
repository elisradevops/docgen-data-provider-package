import { fetchSettledBounded } from '../../utils/boundedFetch';

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
    expect(maxInFlight).toBe(3);
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
