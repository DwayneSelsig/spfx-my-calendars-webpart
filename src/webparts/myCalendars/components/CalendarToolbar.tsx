import * as React from 'react';
import { DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Dropdown, type IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import type { CalendarViewType } from '../models/ICalendarSettings';
import { getCalendarLabels } from './views/calendarLabels';
import { localDateKey, parseLocalDateKey } from './views/calendarUtils';

interface ICalendarViewOptionData {
  iconName: string;
}

const compactViewSelector = '@media screen and (max-width: 479px)';

const toolbarClassNames = mergeStyleSets({
  viewDropdown: {
    width: 128,
    flexShrink: 0,
    [compactViewSelector]: {
      width: 52
    }
  },
  viewOptionContent: {
    display: 'inline-flex',
    alignItems: 'center',
    minWidth: 0
  },
  viewOptionIcon: {
    flexShrink: 0,
    fontSize: 16
  },
  selectedViewText: {
    marginLeft: 8,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    [compactViewSelector]: {
      display: 'none'
    }
  },
  menuViewText: {
    marginLeft: 8
  }
});

const getViewIconName = (option: IDropdownOption): string =>
  (option.data as ICalendarViewOptionData | undefined)?.iconName || 'Calendar';

const renderViewOption = (option?: IDropdownOption): JSX.Element => {
  if (!option) return <React.Fragment />;

  return (
    <span className={toolbarClassNames.viewOptionContent}>
      <Icon className={toolbarClassNames.viewOptionIcon} iconName={getViewIconName(option)} aria-hidden="true" />
      <span className={toolbarClassNames.menuViewText}>{option.text}</span>
    </span>
  );
};

const renderSelectedView = (options?: IDropdownOption[]): JSX.Element => {
  const option = options?.[0];
  if (!option) return <React.Fragment />;

  return (
    <span className={toolbarClassNames.viewOptionContent}>
      <Icon className={toolbarClassNames.viewOptionIcon} iconName={getViewIconName(option)} aria-hidden="true" />
      <span className={toolbarClassNames.selectedViewText}>{option.text}</span>
    </span>
  );
};

export interface ICalendarToolbarProps {
  currentDate: Date;
  currentView: CalendarViewType;
  dateRangeText: string;
  onToday: () => void;
  onNavigate: (direction: 'prev' | 'next') => void;
  onDateChange: (date: Date) => void;
  onViewChange: (view: CalendarViewType) => void;
}

export const CalendarToolbar: React.FC<ICalendarToolbarProps> = ({ currentDate, currentView, dateRangeText, onToday, onNavigate, onDateChange, onViewChange }) => {
  const labels = getCalendarLabels();
  const viewOptions: IDropdownOption<ICalendarViewOptionData>[] = [
    { key: 'day', text: labels.day, data: { iconName: 'CalendarAgenda' } },
    { key: 'week', text: labels.week, data: { iconName: 'CalendarWeek' } },
    { key: 'month', text: labels.month, data: { iconName: 'Calendar' } }
  ];
  const selectedViewLabel = viewOptions.find(option => option.key === currentView)?.text || labels.view;

  return (
    <nav aria-label={labels.navigation} style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 6, padding: '8px 0', width: '100%' }}>
      <DefaultButton text={labels.today} onClick={onToday} />
      <IconButton iconProps={{ iconName: 'ChevronLeft' }} ariaLabel={labels.previous} title={labels.previous} onClick={() => onNavigate('prev')} />
      <IconButton iconProps={{ iconName: 'ChevronRight' }} ariaLabel={labels.next} title={labels.next} onClick={() => onNavigate('next')} />
      <label style={{ display: 'inline-flex', alignItems: 'center' }}>
        <span style={{ position: 'absolute', width: 1, height: 1, overflow: 'hidden', clip: 'rect(0 0 0 0)' }}>{dateRangeText}</span>
        <input
          type="date"
          value={localDateKey(currentDate)}
          aria-label={dateRangeText}
          onChange={event => {
            const parsed = parseLocalDateKey(event.currentTarget.value);
            if (parsed) onDateChange(parsed);
          }}
          style={{ height: 32, boxSizing: 'border-box', border: '1px solid var(--neutralTertiary, #a19f9d)', borderRadius: 2, padding: '0 8px', background: 'var(--white, #fff)', color: 'var(--neutralPrimary, #323130)' }}
        />
      </label>
      <strong style={{ flex: '1 1 180px', minWidth: 140, padding: '0 6px', textTransform: 'capitalize' }}>{dateRangeText}</strong>
      <Dropdown
        ariaLabel={`${labels.view}: ${selectedViewLabel}`}
        className={toolbarClassNames.viewDropdown}
        selectedKey={currentView}
        options={viewOptions}
        onRenderOption={renderViewOption}
        onRenderTitle={renderSelectedView}
        onChange={(_, option) => {
          if (option) onViewChange(option.key as CalendarViewType);
        }}
        styles={{ title: { height: 32, lineHeight: 30 }, dropdown: { height: 32 } }}
      />
    </nav>
  );
};
