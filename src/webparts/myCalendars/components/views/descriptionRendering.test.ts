import { sanitizeCalendarEventHtml } from './descriptionRendering';

describe('sanitizeCalendarEventHtml', () => {
  it('keeps rich formatting and hardens safe links', () => {
    const html = sanitizeCalendarEventHtml('<p><strong>Important</strong></p><table><tbody><tr><td>Cell</td></tr></tbody></table><a href="https://example.com/path">Open</a>');
    const container = document.createElement('div');
    container.innerHTML = html;

    expect(container.querySelector('strong')?.textContent).toBe('Important');
    expect(container.querySelector('td')?.textContent).toBe('Cell');
    expect(container.querySelector('a')?.getAttribute('href')).toBe('https://example.com/path');
    expect(container.querySelector('a')?.getAttribute('target')).toBe('_blank');
    expect(container.querySelector('a')?.getAttribute('rel')).toBe('noopener noreferrer');
  });

  it('removes active content, inline handlers, styles, and unsafe links', () => {
    const html = sanitizeCalendarEventHtml('<script>alert(1)</script><p style="color:red" onclick="alert(2)">Text</p><a href="javascript:alert(3)">Unsafe</a><iframe src="https://example.com"></iframe>');
    const container = document.createElement('div');
    container.innerHTML = html;

    expect(container.querySelector('script')).toBeNull();
    expect(container.querySelector('iframe')).toBeNull();
    expect(container.querySelector('p')?.hasAttribute('style')).toBe(false);
    expect(container.querySelector('p')?.hasAttribute('onclick')).toBe(false);
    expect(container.querySelector('a')?.hasAttribute('href')).toBe(false);
  });
});
