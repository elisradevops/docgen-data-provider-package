// Guards the run context (runId / docType / project) that every log record is stamped from.
// Prod showed data-provider records with a run id but no project / doc type; these tests pin
// down what must hold in this package for a record to be attributed to the right run.
import * as fs from 'fs';
import * as path from 'path';
import * as winston from 'winston';
import { DiagnosticsTransport, withRunContext } from '../../utils/logger';
import { installLogSink, LogSink, DiagnosticEvent } from '../../utils/logSink';
import { runContextStore } from '../../utils/runContext';

const pLimit = require('p-limit');
const pause = (ms: number) => new Promise((r) => setTimeout(r, ms));

function makeLogger() {
  const events: DiagnosticEvent[] = [];
  installLogSink({ push: (e: DiagnosticEvent) => events.push(e), flush: async () => undefined } as LogSink);
  const logger = winston.createLogger({
    level: 'debug',
    format: winston.format.combine(withRunContext(), winston.format.json()),
    transports: [new DiagnosticsTransport()],
  });
  return { logger, events };
}

describe('run context reaches log events from the places the data provider logs from', () => {
  test('from tasks queued behind a per-instance p-limit, and from a retry timer', async () => {
    const { logger, events } = makeLogger();
    const limit = pLimit(2); // per request, like the providers' `private limit = pLimit(10)` fields
    await runContextStore.run({ runId: 'run-1', docType: 'STD', project: 'MEWP' }, async () => {
      await Promise.all(
        [1, 2, 3, 4, 5, 6].map((i) =>
          limit(async () => {
            await pause(3);
            logger.error(`queued task ${i}`);
          })
        )
      );
      await new Promise<void>((resolve) =>
        setTimeout(() => {
          logger.warn('Request failed. Retrying');
          resolve();
        }, 5)
      );
    });
    expect(events).toHaveLength(7);
    events.forEach((e) => expect(e).toMatchObject({ runId: 'run-1', docType: 'STD', project: 'MEWP' }));
  });

  test('concurrent runs, each with its own limiter, never see each other\'s context', async () => {
    const { logger, events } = makeLogger();
    await Promise.all(
      ['A', 'B', 'C', 'D'].map((id) =>
        runContextStore.run({ runId: `run-${id}`, docType: id === 'A' ? 'STD' : 'SVD', project: `P-${id}` }, async () => {
          const limit = pLimit(2);
          await Promise.all([1, 2, 3, 4].map((i) => limit(async () => { await pause(i); logger.error(`${id}:${i}`); })));
        })
      )
    );
    expect(events).toHaveLength(16);
    events.forEach((e) => {
      const id = e.message.split(':')[0];
      expect(e).toMatchObject({ runId: `run-${id}`, project: `P-${id}` });
    });
  });
});

// A p-limit queue shared by two concurrent runs resumes a queued task from whichever task last
// finished — i.e. in the OTHER run's async context — so its log lines are attributed to the wrong
// run (measured: half the tasks of two interleaved runs). Per-instance limiters are safe because
// a provider instance serves one request. If you need a limiter, make it a per-request/instance
// field; do not hoist it to module or class-static scope.
describe('no process-wide p-limit queue', () => {
  const walk = (dir: string): string[] =>
    fs.readdirSync(dir, { withFileTypes: true }).flatMap((d) => {
      const p = path.join(dir, d.name);
      if (d.isDirectory()) return d.name === 'tests' ? [] : walk(p);
      return p.endsWith('.ts') ? [p] : [];
    });

  test('no pLimit(...) at module scope or in a static member', () => {
    const offenders: string[] = [];
    for (const file of walk(path.resolve(__dirname, '../..'))) {
      fs.readFileSync(file, 'utf8')
        .split('\n')
        .forEach((line, i) => {
          const moduleScope = /^(export\s+)?(const|let|var)\s+\w+\s*=\s*pLimit\(/.test(line);
          const staticMember = /\bstatic\b[^=]*=\s*pLimit\(/.test(line);
          if (moduleScope || staticMember) offenders.push(`${path.relative(process.cwd(), file)}:${i + 1}: ${line.trim()}`);
        });
    }
    expect(offenders).toEqual([]);
  });
});
