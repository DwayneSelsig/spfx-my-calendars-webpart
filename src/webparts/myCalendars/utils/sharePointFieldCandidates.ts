import * as strings from 'MyCalendarsWebPartStrings';

export type SharePointFieldKind = 'title' | 'start' | 'end';

const DEFAULT_FIELD_CANDIDATES: Record<SharePointFieldKind, string[]> = {
  title: ['Title', 'Subject', 'Event Title'],
  start: ['Start Time', 'Start', 'Start Date', 'StartDate', 'StartDateTime', 'EventDate', 'Starttijd', 'Begindatum', 'Begin', 'Startdatum'],
  end: ['End Time', 'End', 'End Date', 'EndDate', 'EndDateTime', 'Eindtijd', 'Einddatum', 'Einde']
};

export function parseFieldCandidates(value: string | undefined): string[] {
  return (value || '')
    .split(';')
    .map(candidate => candidate.trim())
    .filter(Boolean);
}

export function normalizeFieldCandidate(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, '');
}

export function getFieldCandidates(field: SharePointFieldKind): string[] {
  const localized = parseFieldCandidates(
    field === 'title' ? strings.FieldTitleCandidates : field === 'start' ? strings.FieldStartCandidates : strings.FieldEndCandidates
  );
  const seen = new Set<string>();

  return [...localized, ...DEFAULT_FIELD_CANDIDATES[field]].filter(candidate => {
    const normalized = normalizeFieldCandidate(candidate);
    if (!normalized || seen.has(normalized)) return false;
    seen.add(normalized);
    return true;
  });
}

export function findBestMatchingFieldKey(options: Array<{ key: string | number; text?: string }>, candidates: string[]): string | number | undefined {
  const normalizedCandidates = candidates.map(normalizeFieldCandidate).filter(Boolean);
  const normalizedOptions = options.map(option => ({
    option,
    text: normalizeFieldCandidate(String(option.text || '')),
    key: normalizeFieldCandidate(String(option.key || ''))
  }));

  for (const candidate of normalizedCandidates) {
    const exactMatch = normalizedOptions.find(entry => entry.text === candidate || entry.key === candidate);
    if (exactMatch) return exactMatch.option.key;
  }

  for (const candidate of normalizedCandidates) {
    const partialMatch = normalizedOptions.find(entry => entry.text.includes(candidate) || entry.key.includes(candidate));
    if (partialMatch) return partialMatch.option.key;
  }

  return undefined;
}
