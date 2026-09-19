import {
  findBestMatchingFieldKey,
  getFieldCandidates,
  parseFieldCandidates
} from './sharePointFieldCandidates';

describe('SharePoint field candidate parsing', () => {
  it('uses semicolons, trims values, and removes empty entries', () => {
    expect(parseFieldCandidates(' Titel;  Onderwerp ; ;Titel ')).toEqual(['Titel', 'Onderwerp', 'Titel']);
    expect(parseFieldCandidates('Start Time;Start;Start Date')).toEqual(['Start Time', 'Start', 'Start Date']);
  });

  it('combines localized candidates with supported defaults without duplicates', () => {
    const candidates = getFieldCandidates('start');
    expect(candidates).toContain('Start Time');
    expect(candidates).toContain('Starttijd');
    expect(candidates.filter(value => value.toLowerCase() === 'start').length).toBe(1);
  });

  it('matches localized display names and internal field names consistently', () => {
    const options = [
      { key: 'StartDate', text: 'Begindatum' },
      { key: 'EventDate', text: 'Event Date' }
    ];

    expect(findBestMatchingFieldKey(options, parseFieldCandidates('Start Time;Begindatum'))).toBe('StartDate');
    expect(findBestMatchingFieldKey(options, parseFieldCandidates('EventDate;Einddatum'))).toBe('EventDate');
  });
});
