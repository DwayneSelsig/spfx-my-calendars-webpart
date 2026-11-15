/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Shared host is unmounted after each test. */
jest.mock('MyCalendarsWebPartStrings', () => new Proxy({ __esModule: true }, {
  get: (_target, key) => key === '__esModule' ? true : String(key)
}));
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { AudienceSelection, type IAudienceSelectionProps } from './AudienceSelection';
import type { AudienceService } from '../services/AudienceService';

describe('audience selection', () => {
  let host: HTMLDivElement;
  let props: IAudienceSelectionProps;
  let count: jest.Mock;
  function render(update: Partial<IAudienceSelectionProps> = {}): void {
    props = { ...props, ...update };
    act(() => { ReactDOM.render(<AudienceSelection {...props} />, host); });
  }
  function button(label: string): HTMLButtonElement {
    const found = Array.from(host.querySelectorAll('button')).find(item => item.textContent === label);
    if (!found) throw new Error('Missing button ' + label);
    return found;
  }
  function select(name: string): void {
    const row = Array.from(host.querySelectorAll('[role="row"]')).find(item => item.textContent?.indexOf(name) !== -1);
    if (!row) throw new Error('Missing row ' + name);
    act(() => { Simulate.click(row); });
  }
  beforeEach(() => {
    host = document.createElement('div'); document.body.appendChild(host);
    count = jest.fn(async (ids: string[]) => {
      const result: Record<string, number> = {};
      ids.forEach(id => { result[id] = id === 'g0' ? 0 : 7; });
      return result;
    });
    props = {
      groups: Array.from({ length: 12 }, (_, index) => ({ id: `g${index}`, displayName: `Group ${index}!`, groupType: 'security' })),
      loading: false, loaded: true, search: '', everyone: false, single: true, selected: {},
      service: { getAudienceMemberCounts: count } as unknown as AudienceService, actionLabel: 'Apply',
      onSearchChange: jest.fn(search => render({ search })),
      onEveryoneChange: jest.fn(everyone => render({ everyone, selected: {} })),
      onSelectionChange: jest.fn(selected => render({ selected })), onApply: jest.fn()
    };
  });
  afterEach(() => { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); jest.restoreAllMocks(); });

  it('counts only the current page, reuses successful counts, and preserves a single selection across pages', async () => {
    await act(async () => { render(); });
    expect(count.mock.calls[0][0]).toHaveLength(10);
    expect(host.textContent).toContain('0');
    const firstRow = Array.from(host.querySelectorAll('[role="row"]')).find(row => row.textContent?.indexOf('Group 0!') !== -1) as HTMLElement;
    act(() => { firstRow.focus(); Simulate.keyDown(firstRow, { key: ' ' }); });
    expect(props.selected).toEqual({ g0: 'Group 0!' });
    await act(async () => { Simulate.click(button('NextLabel')); });
    expect(count.mock.calls[1][0]).toEqual(['g10', 'g11']);
    expect(props.selected).toEqual({ g0: 'Group 0!' });
    select('Group 10!');
    expect(props.selected).toEqual({ g10: 'Group 10!' });
    await act(async () => { Simulate.click(button('AudiencePreviousPageLabel')); });
    expect(count).toHaveBeenCalledTimes(2);
  });

  it('keeps selections on other pages when editing multiple groups', async () => {
    await act(async () => { render({ single: false, selected: { g11: 'Group 11!' } }); });
    select('Group 0!');
    select('Group 1!');
    expect(props.selected).toEqual({ g11: 'Group 11!', g0: 'Group 0!', g1: 'Group 1!' });
  });

  it('requires a selection, and Everyone clears and hides group selection', async () => {
    await act(async () => { render(); });
    expect(button('Apply').disabled).toBe(true);
    select('Group 0!');
    expect(button('Apply').disabled).toBe(false);
    const everyone = host.querySelector('input[value="everyone"]') as HTMLInputElement;
    act(() => { everyone.click(); });
    expect(props.everyone).toBe(true);
    expect(props.selected).toEqual({});
    expect(host.querySelector('[role="grid"]')).toBeNull();
    expect(button('Apply').disabled).toBe(false);
  });

  it('filters immediately by case-insensitive substring across all pages and keeps selection', async () => {
    await act(async () => { render({ selected: { g0: 'Group 0!' } }); });
    await act(async () => { Simulate.click(button('NextLabel')); });
    const input = host.querySelector('input[type="search"]') as HTMLInputElement;
    await act(async () => { Simulate.change(input, { target: { value: 'OUP 1' } as unknown as EventTarget }); });
    expect(host.querySelector('[role="grid"]')?.textContent).toContain('Group 1!');
    expect(host.querySelector('[role="grid"]')?.textContent).toContain('Group 10!');
    expect(host.querySelector('[role="grid"]')?.textContent).not.toContain('Group 0!');
    expect(props.groups).toHaveLength(12);
    expect(props.selected).toEqual({ g0: 'Group 0!' });
    await act(async () => { Simulate.change(input, { target: { value: 'missing' } as unknown as EventTarget }); });
    expect(host.textContent).toContain('NoSecurityGroupsLabel');
    expect(props.selected).toEqual({ g0: 'Group 0!' });
    await act(async () => { Simulate.change(input, { target: { value: '' } as unknown as EventTarget }); });
    expect(host.querySelector('[role="grid"]')?.textContent).toContain('Group 0!');
    expect(host.querySelector('[role="grid"]')?.textContent).not.toContain('Group 10!');
    expect(host.textContent).not.toContain('NoSecurityGroupsLabel');
    expect(count).toHaveBeenCalledTimes(2);
    render({ error: 'Discovery failed', loaded: false });
    expect(host.querySelector('[role="alert"]')?.textContent).toBe('Discovery failed');
    expect(button('Apply').disabled).toBe(false);
  });

  it('ignores obsolete count results and renders failures as unavailable without blocking selection', async () => {
    let finish!: (result: Record<string, number | undefined>) => void;
    count.mockImplementationOnce(() => new Promise(resolve => { finish = resolve; }));
    render();
    await act(async () => { render({ groups: [{ id: 'new', displayName: 'New group', groupType: 'security' }] }); });
    await act(async () => { finish({ g0: 123456789 }); });
    expect(host.textContent).not.toContain('123456789');
    count.mockResolvedValueOnce({ unavailable: undefined });
    await act(async () => { render({ groups: [{ id: 'unavailable', displayName: 'Unavailable group', groupType: 'security' }] }); });
    expect(host.textContent).toContain('AudienceCountUnavailableLabel');
    select('Unavailable group');
    expect(button('Apply').disabled).toBe(false);
  });
});
