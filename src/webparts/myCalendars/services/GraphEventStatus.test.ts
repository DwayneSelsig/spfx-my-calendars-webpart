import { normalizeAvailabilityStatus, normalizeResponseStatus } from './GraphEventStatus';

describe('GraphEventStatus', () => {
  it('normalizes missing and future availability values safely', () => {
    expect(normalizeAvailabilityStatus(undefined)).toBeUndefined();
    expect(normalizeAvailabilityStatus('futureValue')).toBe('unknown');
  });

  it('omits missing and future response values', () => {
    expect(normalizeResponseStatus(undefined)).toBeUndefined();
    expect(normalizeResponseStatus('futureValue')).toBeUndefined();
  });
});
