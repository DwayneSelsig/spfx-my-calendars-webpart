export type LocalizationValue = string | number;

export function formatLocalizedString(template: string, ...values: LocalizationValue[]): string {
  let result = template;
  values.forEach((value, index) => {
    result = result.split(`{${index}}`).join(String(value));
  });
  return result;
}
