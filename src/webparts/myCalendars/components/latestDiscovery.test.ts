import { LatestDiscovery } from './latestDiscovery';

function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void; reject: (error: unknown) => void } {
  let complete!: (value: T) => void;
  let fail!: (error: unknown) => void;
  const promise = new Promise<T>((resolve, reject) => { complete = resolve; fail = reject; });
  return { promise, resolve: complete, reject: fail };
}

describe('settings discovery lifecycle', () => {
  it('clears stale results and finishes on failures', async () => {
    const discovery = new LatestDiscovery();
    let items = ['old'];
    let loading = false;
    let error: unknown;
    const callbacks = {
      start: () => { loading = true; items = []; error = undefined; },
      success: (value: string[]) => { items = value; },
      error: (value: unknown) => { error = value; },
      finish: () => { loading = false; }
    };
    await discovery.run(async () => ['new'], callbacks);
    expect(items).toEqual(['new']);
    expect(loading).toBe(false);
    await discovery.run(async () => { throw new Error('403'); }, callbacks);
    expect(items).toEqual([]);
    expect(error).toEqual(new Error('403'));
    expect(loading).toBe(false);
    await discovery.run(async () => [], callbacks);
    expect(error).toBeUndefined();
    expect(items).toEqual([]);
  });

  it('ignores a stale success after the newer request fails, including its finalizer', async () => {
    const discovery = new LatestDiscovery();
    const old = deferred<string[]>();
    const success = jest.fn();
    const error = jest.fn();
    const finish = jest.fn();
    const callbacks = { start: jest.fn(), success, error, finish };
    const pending = discovery.run(() => old.promise, callbacks);
    await discovery.run(async () => { throw new Error('404'); }, callbacks);
    old.resolve(['obsolete']);
    await pending;
    expect(success).not.toHaveBeenCalled();
    expect(error).toHaveBeenCalledTimes(1);
    expect(finish).toHaveBeenCalledTimes(1);
  });

  it('does not overwrite an active request loading state when an older request rejects', async () => {
    const discovery = new LatestDiscovery();
    const first = deferred<string[]>();
    const second = deferred<string[]>();
    const callbacks = { start: jest.fn(), success: jest.fn(), error: jest.fn(), finish: jest.fn() };
    const a = discovery.run(() => first.promise, callbacks);
    const b = discovery.run(() => second.promise, callbacks);
    first.reject(new Error('old failure'));
    await a;
    expect(callbacks.finish).not.toHaveBeenCalled();
    expect(callbacks.error).not.toHaveBeenCalled();
    second.resolve([]);
    await b;
    expect(callbacks.finish).toHaveBeenCalledTimes(1);
  });

  it('ignores callbacks after changed input, closing or unmounting', async () => {
    const discovery = new LatestDiscovery();
    const request = deferred<string[]>();
    const callbacks = { start: jest.fn(), success: jest.fn(), error: jest.fn(), finish: jest.fn() };
    const pending = discovery.run(() => request.promise, callbacks);
    discovery.invalidate();
    request.reject(new Error('cancelled dialog'));
    await pending;
    expect(callbacks.success).not.toHaveBeenCalled();
    expect(callbacks.error).not.toHaveBeenCalled();
    expect(callbacks.finish).not.toHaveBeenCalled();
  });
});
