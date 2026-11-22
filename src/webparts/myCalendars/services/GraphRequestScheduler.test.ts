import { GraphRequestCancelled, GraphRequestScheduler, GraphRequestScope, normalizeGraphConcurrency } from './GraphRequestScheduler';

function deferred(): { promise: Promise<void>; resolve: () => void } {
  let finish!: () => void;
  const promise = new Promise<void>(resolve => { finish = resolve; });
  return { promise, resolve: finish };
}

function harness(limit?: number): {
  scheduler: GraphRequestScheduler;
  started: string[];
  add: (name: string, priority?: boolean | GraphRequestScope) => { done: Promise<void>; resolve: () => void };
  finish: () => Promise<void>;
} {
  const scheduler = new GraphRequestScheduler(limit);
  const started: string[] = [];
  const entries: Array<{ done: Promise<void>; resolve: () => void }> = [];
  return {
    scheduler, started,
    add: (name, priority) => {
      const gate = deferred();
      const done = scheduler.run(() => { started.push(name); return gate.promise; }, priority);
      entries.push({ done, resolve: gate.resolve });
      return { done, resolve: gate.resolve };
    },
    finish: async () => {
      entries.forEach(entry => entry.resolve());
      await Promise.all(entries.map(entry => entry.done.catch(() => undefined)));
    }
  };
}

describe('Graph request scheduler', () => {
  it.each([[undefined, 10], [0, 2], [2, 2], [20, 20], [100, 20], [2.5, 10], [NaN, 10], [Infinity, 10], ['12', 10]])(
    'normalizes %p to %p', (value, expected) => expect(normalizeGraphConcurrency(value)).toBe(expected)
  );

  it.each([5, 10, 15, 19, 20])('reserves the first urgent start at limit %p and only opens one burst', async limit => {
    const work = harness(limit);
    for (let index = 0; index <= limit; index++) work.add(`normal ${index}`);
    const normal = limit === 20 ? 19 : limit;
    expect(work.scheduler.counts.activeNormal).toBe(normal);
    work.add('first high', true);
    work.add('second high', true);
    work.add('third high', true);
    expect(work.started).toContain('first high');
    expect(work.started).not.toContain('second high');
    expect(work.scheduler.counts).toMatchObject({ activeNormal: normal, activeHigh: 1, queuedHigh: 2 });
    await work.finish();
    expect(work.started.indexOf('second high')).toBeLessThan(work.started.indexOf('third high'));
    expect(new Set(work.started).size).toBe(work.started.length);
  });

  it.each([[2, 1], [10, 1], [11, 2], [20, 2]])('gives normal loads their minimum under sustained high demand at %p', async (limit, minimum) => {
    const work = harness(limit);
    const high = Array.from({ length: limit + 4 }, (_, index) => work.add(`high ${index}`, true));
    const normal = Array.from({ length: 4 }, (_, index) => work.add(`normal ${index}`));
    expect(work.scheduler.counts.activeHigh).toBe(limit);
    for (let index = 0; index < minimum + 1; index++) {
      high[index].resolve();
      await high[index].done;
    }
    expect(work.scheduler.counts.activeNormal).toBe(minimum);
    expect(work.started.slice(limit, limit + minimum)).toEqual(Array.from({ length: minimum }, (_, index) => `normal ${index}`));
    expect(work.started[limit + minimum]).toBe(`high ${limit}`);
    normal[0].resolve();
    await normal[0].done;
    expect(work.started).toContain(`normal ${minimum}`);
    await work.finish();
  });

  it('promotes pending work, keeps FIFO in each queue, and does not burst mixed loads', async () => {
    const work = harness(2);
    const normal = work.add('normal active');
    const high = work.add('high active', true);
    const scope = new GraphRequestScope();
    work.add('normal first');
    work.add('normal second');
    work.add('high first', true);
    work.add('promoted', scope);
    scope.setHigh(true);
    expect(work.started).toEqual(['normal active', 'high active']);
    high.resolve();
    await high.done;
    expect(work.started).toHaveLength(3);
    expect(work.started[2]).toBe('high first');
    normal.resolve();
    await normal.done;
    expect(work.started[3]).toBe('normal first');
    await work.finish();
    expect(work.started.indexOf('normal first')).toBeLessThan(work.started.indexOf('normal second'));
    expect(work.started.indexOf('high first')).toBeLessThan(work.started.indexOf('promoted'));
  });

  it('resizes active work without interrupting it and removes cancelled queued work', async () => {
    const work = harness();
    const active = Array.from({ length: 10 }, (_, index) => work.add(`active ${index}`));
    work.scheduler.setLimit(2);
    const scope = new GraphRequestScope();
    scope.setHigh(true);
    const cancelled = work.add('cancelled', scope);
    const rejection = expect(cancelled.done).rejects.toBeInstanceOf(GraphRequestCancelled);
    scope.cancel();
    await rejection;
    work.add('waiting');
    active[0].resolve();
    await active[0].done;
    expect(work.started).toHaveLength(10);
    work.scheduler.setLimit(11);
    expect(work.started[10]).toBe('waiting');
    await work.finish();
    expect(work.scheduler.counts).toMatchObject({ activeNormal: 0, activeHigh: 0, queuedNormal: 0, queuedHigh: 0 });
  });

  it('releases slots after synchronous throws, rejected promises and success', async () => {
    const scheduler = new GraphRequestScheduler(2);
    await expect(scheduler.run(() => { throw new Error('synchronous'); })).rejects.toThrow('synchronous');
    await expect(scheduler.run(async () => { throw new Error('rejected'); })).rejects.toThrow('rejected');
    await expect(scheduler.run(async () => 'success')).resolves.toBe('success');
    expect(scheduler.counts.activeNormal).toBe(0);
  });
});
