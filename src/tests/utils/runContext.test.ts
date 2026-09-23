import * as winston from 'winston';
import Transport = require('winston-transport');
import { withRunContext } from '../../utils/logger';
import { runContextStore } from '../../utils/runContext';

class CaptureTransport extends Transport {
  lines: Record<string, unknown>[] = [];
  log(info: Record<string, unknown>, callback: () => void) {
    this.lines.push(JSON.parse((info as any)[Symbol.for('message')] ?? JSON.stringify(info)));
    callback();
  }
}
function makeTestLogger() {
  const capture = new CaptureTransport();
  const logger = winston.createLogger({
    level: 'silly',
    format: winston.format.combine(withRunContext(), winston.format.json()),
    transports: [capture],
  });
  return { logger, capture };
}

describe('withRunContext', () => {
  test('is a no-op when the store was never populated', () => {
    const { logger, capture } = makeTestLogger();
    logger.info('outside any run');
    expect(capture.lines[0].runId).toBeUndefined();
  });

  test('stamps runId onto every log emitted inside store.run(...)', () => {
    const { logger, capture } = makeTestLogger();
    runContextStore.run({ runId: 'run-123' }, () => {
      logger.info('inside the run');
    });
    logger.info('outside again');
    expect(capture.lines[0].runId).toBe('run-123');
    expect(capture.lines[1].runId).toBeUndefined();
  });

  test('async work started inside the run keeps the runId across the await boundary', async () => {
    const { logger, capture } = makeTestLogger();
    await runContextStore.run({ runId: 'run-async' }, async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      logger.info('after await');
    });
    expect(capture.lines[0].runId).toBe('run-async');
  });
});
