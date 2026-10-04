import {
  sanitizeUrl,
  summarizeBody,
  summarizeResponse,
  buildAdoRequestContext,
  annotateAdoError,
  maskWiqlLiterals,
} from '../../helpers/requestContext';

describe('maskWiqlLiterals', () => {
  it('replaces every quoted literal with its length', () => {
    expect(maskWiqlLiterals("a = 'xy' and b = ''")).toBe("a = '[2 chars]' and b = '[0 chars]'");
  });
});

describe('sanitizeUrl', () => {
  it('returns undefined for a missing or empty url', () => {
    expect(sanitizeUrl(undefined)).toBeUndefined();
    expect(sanitizeUrl('')).toBeUndefined();
  });

  it('keeps a normal ADO url intact', () => {
    const url = 'https://dev.azure.com/org/proj/_apis/wit/queries/abc?$depth=1&api-version=7.0';
    expect(sanitizeUrl(url)).toBe(url);
  });

  it('strips credentials embedded in the url', () => {
    const out = sanitizeUrl('https://user:s3cret@dev.azure.com/org/_apis/projects');
    expect(out).not.toContain('s3cret');
    expect(out).not.toContain('user');
    expect(out).toContain('dev.azure.com/org/_apis/projects');
  });

  it('masks sensitive query parameter values but keeps ordinary ones', () => {
    const out = sanitizeUrl('https://host/api?access_token=abc123&top=10&sig=zzz&keywords=release')!;
    expect(out).not.toContain('abc123');
    expect(out).not.toContain('zzz');
    expect(out).toContain('top=10');
    expect(out).toContain('keywords=release');
  });

  it('leaves non-sensitive parameters byte-identical when masking another', () => {
    expect(sanitizeUrl('https://h/api?$filter=a%20b&token=x&q=a+b')).toBe('https://h/api?$filter=a%20b&token=[REDACTED]&q=a+b');
  });

  it('passes a non-absolute url through, bounded', () => {
    expect(sanitizeUrl('/relative/path')).toBe('/relative/path');
    expect(sanitizeUrl('x'.repeat(5000))!.length).toBeLessThanOrEqual(1001);
  });
});

describe('summarizeBody', () => {
  it('never describes a GET body', () => {
    expect(summarizeBody({ query: 'SELECT 1' }, 'get')).toBeUndefined();
    expect(summarizeBody({ query: 'SELECT 1' }, 'HEAD')).toBeUndefined();
  });

  it('returns undefined for an empty body — the default `{}` callers pass', () => {
    expect(summarizeBody({}, 'post')).toBeUndefined();
    expect(summarizeBody([], 'post')).toBeUndefined();
    expect(summarizeBody(null, 'post')).toBeUndefined();
    expect(summarizeBody('', 'post')).toBeUndefined();
  });

  it('keeps a WIQL-style body so the failing query is visible', () => {
    const out = summarizeBody({ query: "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = 'P'" }, 'post')!;
    expect(out).toContain('SELECT [System.Id] FROM WorkItems');
    expect(out).toContain("[System.TeamProject] = '[1 chars]'");
    expect(out).not.toContain("= 'P'");
  });

  it('masks WIQL literals (incl. escaped quotes) but keeps fields, numbers and macros', () => {
    const out = summarizeBody({ query: "SELECT [Id] FROM WorkItems WHERE [Title] = 'it''s secret' AND [Assigned] = @Me AND [Id] > 42" }, 'post')!;
    expect(out).not.toContain('secret');
    expect(out).toContain("[Title] = '[12 chars]'");
    expect(out).toContain('@Me');
    expect(out).toContain('> 42');
  });

  it('keeps an id-list body', () => {
    expect(summarizeBody({ ids: [1, 2, 3], fields: ['System.Title'] }, 'post')).toBe(
      '{"ids":[1,2,3],"fields":["System.Title"]}'
    );
  });

  it('redacts credential-looking keys at any depth', () => {
    const out = summarizeBody({ query: 'q', auth: { token: 'tok-123', password: 'pw-456' }, pat: 'pat-789' }, 'post')!;
    expect(out).not.toContain('tok-123');
    expect(out).not.toContain('pw-456');
    expect(out).not.toContain('pat-789');
    expect(out).toContain('[REDACTED]');
    expect(out).toContain('"query":"q"');
  });

  it('JSON-Patch bodies keep op + path only — field values are dropped', () => {
    const patch = [
      { op: 'add', path: '/fields/System.Title', value: 'Highly classified title' },
      { op: 'replace', path: '/fields/System.Description', value: { html: '<p>secret</p>' } },
    ];
    const out = summarizeBody(patch, 'patch')!;
    expect(out).toContain('json-patch (2 ops, values omitted)');
    expect(out).toContain('add /fields/System.Title');
    expect(out).toContain('replace /fields/System.Description');
    expect(out).not.toContain('classified');
    expect(out).not.toContain('secret');
  });

  it('bounds a long string value and the whole body', () => {
    const one = summarizeBody({ query: 'x'.repeat(5000) }, 'post')!;
    expect(one.length).toBeLessThan(700);
    const many = summarizeBody(Object.fromEntries(Array.from({ length: 300 }, (_, i) => [`k${i}`, 'v'.repeat(40)])), 'post')!;
    expect(many.length).toBeLessThanOrEqual(2001);
  });

  it('caps a very long array', () => {
    const out = summarizeBody({ ids: Array.from({ length: 500 }, (_, i) => i) }, 'post')!;
    expect(out).toContain('more');
  });

  it('describes non-JSON bodies by type and size only', () => {
    expect(summarizeBody(Buffer.from('abcdef'), 'post')).toBe('[binary, 6 bytes]');
    expect(summarizeBody('not json at all', 'post')).toBe('[text, 15 chars]');
    expect(summarizeBody(42, 'post')).toBe('[number]');
    expect(summarizeBody({ pipe: () => undefined }, 'post')).toBe('[stream]');
  });

  it('parses a JSON string body and summarizes it like an object', () => {
    expect(summarizeBody('{"query":"q"}', 'post')).toBe('{"query":"q"}');
  });

  it('never throws on a body it cannot serialize', () => {
    expect(summarizeBody({ n: BigInt(1) }, 'post')).toBe('[unserializable body]');
    const circular: Record<string, unknown> = { a: 1 };
    circular.self = circular;
    expect(() => summarizeBody(circular, 'post')).not.toThrow();
  });
});

describe('summarizeResponse', () => {
  it('prefers the server message (ADO puts its explanation there)', () => {
    expect(summarizeResponse({ message: 'TF401232: Work item 5 does not exist', typeName: 'x' })).toBe(
      'TF401232: Work item 5 does not exist'
    );
  });

  it('truncates a text response (e.g. an HTML error page) to 200 chars', () => {
    expect(summarizeResponse('<html>'.repeat(200))!.length).toBeLessThanOrEqual(201);
  });

  it('falls back to a redacted, bounded JSON dump when there is no message', () => {
    const out = summarizeResponse({ code: 1, token: 'leak-me' })!;
    expect(out).not.toContain('leak-me');
    expect(out).toContain('"code":1');
  });

  it('does not walk a stream response', () => {
    expect(summarizeResponse({ pipe: () => undefined, socket: { a: 1 } })).toBe('[stream]');
  });

  it('ignores a missing or binary response', () => {
    expect(summarizeResponse(undefined)).toBeUndefined();
    expect(summarizeResponse(Buffer.from('x'))).toBeUndefined();
  });
});

describe('buildAdoRequestContext', () => {
  it('describes a failed GET with method, status, attempt and the server message', () => {
    const err = { response: { status: 404, data: { message: 'not found' } } };
    expect(buildAdoRequestContext(err, { method: 'get', url: 'https://h/_apis/x' }, 2)).toEqual({
      method: 'GET',
      url: 'https://h/_apis/x',
      status: 404,
      attempt: 2,
      responseExcerpt: 'not found',
    });
  });

  it('includes a summarized body for a POST and omits fields it does not have', () => {
    const ctx = buildAdoRequestContext(new Error('boom'), { method: 'post', url: 'https://h/wiql', data: { query: 'q' } });
    expect(ctx).toEqual({ method: 'POST', url: 'https://h/wiql', requestBody: '{"query":"q"}' });
  });

  it('falls back to the axios error config when the request is not described explicitly', () => {
    const ctx = buildAdoRequestContext({ config: { method: 'put', url: 'https://h/y' } }, {});
    expect(ctx.method).toBe('PUT');
    expect(ctx.url).toBe('https://h/y');
  });
});

describe('annotateAdoError', () => {
  it('attaches the context as an enumerable property winston will merge onto the record', () => {
    const err = annotateAdoError(new Error('x'), { method: 'GET', url: 'https://h/z' });
    expect(Object.keys(err)).toContain('adoRequest');
    expect((err as any).adoRequest.url).toBe('https://h/z');
  });

  it('never throws on a frozen or non-object error', () => {
    expect(() => annotateAdoError(Object.freeze(new Error('x')), { url: 'u' })).not.toThrow();
    expect(() => annotateAdoError('a string' as any, { url: 'u' })).not.toThrow();
    expect(() => annotateAdoError(undefined as any, { url: 'u' })).not.toThrow();
  });
});
