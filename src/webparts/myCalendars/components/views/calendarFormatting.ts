/**
 * Formatting helpers for the current SharePoint culture. A locale is supplied by
 * the web part so date and time output does not depend on the browser language.
 */
export function resolveCalendarLocale(locale?: string): string | undefined {
  if (!locale) return undefined;

  try {
    return Intl.DateTimeFormat.supportedLocalesOf([locale]).length > 0 ? locale : undefined;
  } catch {
    return undefined;
  }
}

export function formatCalendarTime(value: Date, locale?: string): string {
  return new Intl.DateTimeFormat(resolveCalendarLocale(locale), {
    hour: 'numeric',
    minute: '2-digit'
  }).format(value);
}

export function formatCalendarDate(value: Date, options: Intl.DateTimeFormatOptions, locale?: string): string {
  return new Intl.DateTimeFormat(resolveCalendarLocale(locale), options).format(value);
}

export function formatCalendarDateTime(value: Date, locale?: string): string {
  return new Intl.DateTimeFormat(resolveCalendarLocale(locale), {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  }).format(value);
}
