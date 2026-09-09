import DOMPurify from 'dompurify';

const forbiddenTags = [
  'base', 'button', 'embed', 'form', 'iframe', 'input', 'link', 'meta',
  'object', 'option', 'select', 'style', 'svg', 'math', 'textarea'
];

/** Sanitizes source HTML and hardens every retained link for a new tab. */
export function sanitizeCalendarEventHtml(value: string): string {
  const sanitized = DOMPurify.sanitize(value, {
    USE_PROFILES: { html: true },
    FORBID_TAGS: forbiddenTags,
    FORBID_ATTR: ['style']
  });
  const template = document.createElement('template');
  template.innerHTML = sanitized;
  template.content.querySelectorAll('a').forEach(link => {
    link.setAttribute('target', '_blank');
    link.setAttribute('rel', 'noopener noreferrer');
  });
  return template.innerHTML;
}
