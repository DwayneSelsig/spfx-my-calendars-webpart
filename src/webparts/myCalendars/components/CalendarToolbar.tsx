import * as React from 'react';
import {
    ActionButton,
    DefaultButton,
    SearchBox,
    Callout,
    Text,
    IconButton
} from '@fluentui/react';
import { mergeStyles } from '@fluentui/react/lib/Styling';

export interface ICalendarToolbarProps {
    currentDate: Date;
    viewType: 'day' | 'week' | 'month' | 'schedule';
    onToday: () => void;
    onDateChange: (date: Date) => void;
    onNavigate: (direction: 'prev' | 'next') => void;
    onSearch?: (query: string) => void;
    dateRangeText: string;
}

interface ICalendarToolbarState {
    isDatePickerOpen: boolean;
    searchQuery: string;
    pickerMonth: number;
    pickerYear: number;
    selectedWeekStart?: Date;
}

export class CalendarToolbar extends React.Component<ICalendarToolbarProps, ICalendarToolbarState> {
    private datePickerButtonRef = React.createRef<HTMLDivElement>();

    constructor(props: ICalendarToolbarProps) {
        super(props);
        this.state = {
            isDatePickerOpen: false,
            searchQuery: '',
            pickerMonth: props.currentDate.getMonth(),
            pickerYear: props.currentDate.getFullYear(),
            selectedWeekStart: undefined
        };
    }

    private handleSearchChange = (_event?: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        const value = newValue || '';
        this.setState({ searchQuery: value });
        if (this.props.onSearch) {
            this.props.onSearch(value);
        }
    };

    private getWeekStart = (date: Date): Date => {
        const d = new Date(date);
        const day = d.getDay();
        const diff = d.getDate() - day + (day === 0 ? -6 : 1);
        return new Date(d.setDate(diff));
    };

    private selectWeek = (weekStartDate: Date): void => {
        this.props.onDateChange(weekStartDate);
        this.setState({ isDatePickerOpen: false });
    };

    private renderDayPicker = (): React.ReactElement => {
        const { currentDate } = this.props;
        const { pickerMonth, pickerYear } = this.state;

        const monthDate = new Date(pickerYear, pickerMonth, 1);
        const firstDay = monthDate.getDay();
        const daysInMonth = new Date(pickerYear, pickerMonth + 1, 0).getDate();
        const daysInPrevMonth = new Date(pickerYear, pickerMonth, 0).getDate();

        const weeks: Date[][] = [];
        let week: Date[] = [];
        let dayCounter = 1;

        // Fill first week with prev month days
        for (let i = firstDay === 0 ? 6 : firstDay - 1; i > 0; i--) {
            week.unshift(new Date(pickerYear, pickerMonth - 1, daysInPrevMonth - i + 1));
        }

        // Fill rest of month
        while (dayCounter <= daysInMonth) {
            week.push(new Date(pickerYear, pickerMonth, dayCounter));
            dayCounter++;
            if (week.length === 7) {
                weeks.push(week);
                week = [];
            }
        }

        // Fill last week with next month days
        let nextDayCounter = 1;
        while (week.length < 7) {
            week.push(new Date(pickerYear, pickerMonth + 1, nextDayCounter));
            nextDayCounter++;
        }
        weeks.push(week);

        return (
            <div style={{ padding: 16, maxWidth: 320 }}>
                <div style={{ marginBottom: 12, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <IconButton
                        iconProps={{ iconName: 'ChevronUp' }}
                        onClick={() => this.setState({ pickerMonth: pickerMonth - 1 < 0 ? 11 : pickerMonth - 1, pickerYear: pickerMonth - 1 < 0 ? pickerYear - 1 : pickerYear })}
                    />
                    <Text variant="medium" style={{ fontWeight: 600 }}>
                        {monthDate.toLocaleDateString(undefined, { month: 'long', year: 'numeric' })}
                    </Text>
                    <IconButton
                        iconProps={{ iconName: 'ChevronDown' }}
                        onClick={() => this.setState({ pickerMonth: pickerMonth + 1 > 11 ? 0 : pickerMonth + 1, pickerYear: pickerMonth + 1 > 11 ? pickerYear + 1 : pickerYear })}
                    />
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4, marginBottom: 8 }}>
                    {['M', 'T', 'W', 'T', 'F', 'S', 'S'].map(d => (
                        <div key={d} style={{ textAlign: 'center', fontSize: 12, fontWeight: 600, padding: 4 }}>
                            {d}
                        </div>
                    ))}
                </div>

                <div>
                    {weeks.map((weekDays, weekIdx) => (
                        <div key={weekIdx} style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4, marginBottom: 4 }}>
                            {weekDays.map((day, dayIdx) => {
                                const isCurrentDay = day.getDate() === currentDate.getDate() &&
                                    day.getMonth() === currentDate.getMonth() &&
                                    day.getFullYear() === currentDate.getFullYear();
                                const isThisMonth = day.getMonth() === pickerMonth;

                                return (
                                    <button
                                        key={dayIdx}
                                        onClick={() => {
                                            this.props.onDateChange(day);
                                            this.setState({ isDatePickerOpen: false });
                                        }}
                                        style={{
                                            padding: 8,
                                            border: isCurrentDay ? '2px solid var(--themePrimary, #0078d4)' : '1px solid transparent',
                                            borderRadius: 4,
                                            backgroundColor: isCurrentDay ? 'var(--themeLighter, #eef6fc)' : 'transparent',
                                            cursor: 'pointer',
                                            fontSize: 13,
                                            fontWeight: isCurrentDay ? 600 : 400,
                                            color: isThisMonth ? 'var(--neutralPrimary, #323130)' : 'var(--neutralTertiary, #a19f9d)',
                                            transition: 'all 0.2s'
                                        }}
                                        onMouseEnter={(e) => {
                                            if (!isCurrentDay) {
                                                (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'var(--neutralLighter, #f3f2f1)';
                                            }
                                        }}
                                        onMouseLeave={(e) => {
                                            if (!isCurrentDay) {
                                                (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
                                            }
                                        }}
                                    >
                                        {day.getDate()}
                                    </button>
                                );
                            })}
                        </div>
                    ))}
                </div>
            </div>
        );
    };

    private renderWeekPicker = (): React.ReactElement => {
        const { currentDate } = this.props;
        const weekStart = this.getWeekStart(currentDate);
        const { pickerMonth, pickerYear } = this.state;

        const monthDate = new Date(pickerYear, pickerMonth, 1);
        const firstDay = monthDate.getDay();
        const daysInMonth = new Date(pickerYear, pickerMonth + 1, 0).getDate();
        const daysInPrevMonth = new Date(pickerYear, pickerMonth, 0).getDate();

        const weeks: Date[][] = [];
        let week: Date[] = [];
        let dayCounter = 1;

        // Fill first week with prev month days
        for (let i = firstDay === 0 ? 6 : firstDay - 1; i > 0; i--) {
            week.unshift(new Date(pickerYear, pickerMonth - 1, daysInPrevMonth - i + 1));
        }

        // Fill rest of month
        while (dayCounter <= daysInMonth) {
            week.push(new Date(pickerYear, pickerMonth, dayCounter));
            dayCounter++;
            if (week.length === 7) {
                weeks.push(week);
                week = [];
            }
        }

        // Fill last week with next month days
        let nextDayCounter = 1;
        while (week.length < 7) {
            week.push(new Date(pickerYear, pickerMonth + 1, nextDayCounter));
            nextDayCounter++;
        }
        weeks.push(week);

        return (
            <div style={{ padding: 16, maxWidth: 400 }}>
                <div style={{ marginBottom: 12, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <IconButton
                        iconProps={{ iconName: 'ChevronUp' }}
                        onClick={() => this.setState({ pickerMonth: pickerMonth - 1 < 0 ? 11 : pickerMonth - 1, pickerYear: pickerMonth - 1 < 0 ? pickerYear - 1 : pickerYear })}
                    />
                    <Text variant="medium" style={{ fontWeight: 600 }}>
                        {monthDate.toLocaleDateString(undefined, { month: 'long', year: 'numeric' })}
                    </Text>
                    <IconButton
                        iconProps={{ iconName: 'ChevronDown' }}
                        onClick={() => this.setState({ pickerMonth: pickerMonth + 1 > 11 ? 0 : pickerMonth + 1, pickerYear: pickerMonth + 1 > 11 ? pickerYear + 1 : pickerYear })}
                    />
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
                    {['M', 'T', 'W', 'T', 'F', 'S', 'S'].map(d => (
                        <div key={d} style={{ textAlign: 'center', fontSize: 12, fontWeight: 600, padding: 4 }}>
                            {d}
                        </div>
                    ))}
                </div>

                {weeks.map((weekDays, weekIdx) => {
                    const week = this.getWeekStart(weekDays[0]);
                    const isCurrentWeek = week.getTime() === weekStart.getTime();

                    return (
                        <div
                            key={weekIdx}
                            onClick={() => this.selectWeek(week)}
                            style={{
                                display: 'grid',
                                gridTemplateColumns: 'repeat(7, 1fr)',
                                gap: 4,
                                padding: 8,
                                marginBottom: 2,
                                border: isCurrentWeek ? '2px solid var(--neutralPrimary, #323130)' : '2px solid transparent',
                                borderRadius: 4,
                                cursor: 'pointer',
                                backgroundColor: isCurrentWeek ? 'var(--neutralLighter, #f3f2f1)' : 'transparent',
                                transition: 'background-color 0.2s'
                            }}
                            onMouseEnter={(e) => {
                                if (!isCurrentWeek) {
                                    (e.currentTarget as HTMLDivElement).style.backgroundColor = 'var(--neutralLighter, #f3f2f1)';
                                }
                            }}
                            onMouseLeave={(e) => {
                                if (!isCurrentWeek) {
                                    (e.currentTarget as HTMLDivElement).style.backgroundColor = 'transparent';
                                }
                            }}
                        >
                            {weekDays.map((day, dayIdx) => (
                                <div
                                    key={dayIdx}
                                    style={{
                                        textAlign: 'center',
                                        fontSize: 13,
                                        padding: 4,
                                        color: day.getMonth() === pickerMonth ? 'var(--neutralPrimary, #323130)' : 'var(--neutralTertiary, #a19f9d)',
                                    }}
                                >
                                    {day.getDate()}
                                </div>
                            ))}
                        </div>
                    );
                })}
            </div>
        );
    };

    private renderMonthPicker = (): React.ReactElement => {
        const { pickerYear } = this.state;
        const { currentDate } = this.props;
        const currentMonth = currentDate.getMonth();
        const currentYear = currentDate.getFullYear();

        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

        return (
            <div style={{ padding: 16, minWidth: 300 }}>
                <div style={{ marginBottom: 20 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                        <Text variant="large" style={{ fontWeight: 600 }}>
                            {pickerYear}
                        </Text>
                        <div>
                            <IconButton
                                iconProps={{ iconName: 'ChevronUp' }}
                                onClick={() => this.setState({ pickerYear: pickerYear + 1 })}
                            />
                            <IconButton
                                iconProps={{ iconName: 'ChevronDown' }}
                                onClick={() => this.setState({ pickerYear: pickerYear - 1 })}
                            />
                        </div>
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 8 }}>
                        {months.map((month, idx) => (
                            <button
                                key={month}
                                onClick={() => {
                                    const newDate = new Date(pickerYear, idx, 1);
                                    this.props.onDateChange(newDate);
                                    this.setState({ isDatePickerOpen: false, pickerMonth: idx });
                                }}
                                style={{
                                    padding: 10,
                                    border: idx === currentMonth && pickerYear === currentYear ? '2px solid var(--themePrimary, #0078d4)' : '1px solid var(--neutralLight, #edebe9)',
                                    borderRadius: 4,
                                    backgroundColor: idx === currentMonth && pickerYear === currentYear ? 'var(--themeLighter, #eef6fc)' : 'var(--white, #ffffff)',
                                    cursor: 'pointer',
                                    fontSize: 14,
                                    fontWeight: 500,
                                    color: 'var(--neutralPrimary, #323130)',
                                    transition: 'all 0.2s'
                                }}
                                onMouseEnter={(e) => {
                                    if (idx !== currentMonth || pickerYear !== currentYear) {
                                        (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'var(--neutralLighter, #f3f2f1)';
                                    }
                                }}
                                onMouseLeave={(e) => {
                                    if (idx !== currentMonth || pickerYear !== currentYear) {
                                        (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'var(--white, #ffffff)';
                                    }
                                }}
                            >
                                {month}
                            </button>
                        ))}
                    </div>
                </div>
            </div>
        );
    };

    private getDatePickerContent = (): React.ReactElement => {
        const { viewType } = this.props;

        switch (viewType) {
            case 'day':
                return this.renderDayPicker();
            case 'week':
                return this.renderWeekPicker();
            case 'month':
            case 'schedule':
            default:
                return this.renderMonthPicker();
        }
    };

    private getNavigationAriaLabel = (direction: 'prev' | 'next'): string => {
        const navigationLabels = {
            day: direction === 'prev' ? 'Previous day' : 'Next day',
            week: direction === 'prev' ? 'Previous week' : 'Next week',
            month: direction === 'prev' ? 'Previous month' : 'Next month',
            schedule: direction === 'prev' ? 'Previous month' : 'Next month'
        };
        return navigationLabels[this.props.viewType];
    };

    private getSearchPlaceholder = (): string => {
        return 'Search appointments...';
    };

    private getWeekDays = (): Date[] => {
        const { currentDate } = this.props;
        const weekStart = this.getWeekStart(currentDate);
        const days: Date[] = [];
        for (let i = 0; i < 7; i++) {
            const day = new Date(weekStart);
            day.setDate(weekStart.getDate() + i);
            days.push(day);
        }
        return days;
    };

    private isSameDay = (date1: Date, date2: Date): boolean => {
        return date1.getDate() === date2.getDate() &&
            date1.getMonth() === date2.getMonth() &&
            date1.getFullYear() === date2.getFullYear();
    };

    private renderDaySelector = (): React.ReactElement | null => {
        const { viewType, currentDate } = this.props;
        
        if (viewType !== 'day' && viewType !== 'schedule') {
            return null;
        }

        const weekDays = this.getWeekDays();
        const today = new Date();

        return (
            <div style={{
                display: 'flex',
                gap: 8,
                marginTop: 8,
                paddingTop: 8,
                borderTop: '1px solid var(--neutralLight, #edebe9)'
            }}>
                {weekDays.map((day, index) => {
                    const isSelected = this.isSameDay(day, currentDate);
                    const isToday = this.isSameDay(day, today);
                    const dayName = day.toLocaleDateString(undefined, { weekday: 'short' });
                    const dayNum = day.getDate();

                    return (
                        <button
                            key={index}
                            onClick={() => this.props.onDateChange(day)}
                            style={{
                                flex: 1,
                                display: 'flex',
                                flexDirection: 'column',
                                alignItems: 'center',
                                padding: '8px 4px',
                                border: isSelected ? '2px solid var(--themePrimary, #0078d4)' : '1px solid var(--neutralLight, #edebe9)',
                                borderRadius: 6,
                                backgroundColor: isSelected ? 'var(--themeLighter, #eef6fc)' : 'transparent',
                                cursor: 'pointer',
                                transition: 'all 0.2s',
                                minWidth: 50
                            }}
                            onMouseEnter={(e) => {
                                if (!isSelected) {
                                    (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'var(--neutralLighter, #f3f2f1)';
                                }
                            }}
                            onMouseLeave={(e) => {
                                if (!isSelected) {
                                    (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
                                }
                            }}
                        >
                            <div style={{
                                fontSize: 11,
                                fontWeight: 500,
                                color: 'var(--neutralSecondary, #605e5c)',
                                marginBottom: 4
                            }}>
                                {dayName}
                            </div>
                            <div style={{
                                fontSize: 18,
                                fontWeight: isSelected ? 600 : 500,
                                backgroundColor: isToday && !isSelected ? 'var(--themePrimary, #0078d4)' : 'transparent',
                                color: isToday && !isSelected ? 'white' : (isToday ? 'var(--themePrimary, #0078d4)' : 'var(--neutralPrimary, #323130)'),
                                borderRadius: '50%',
                                width: 32,
                                height: 32,
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center'
                            }}>
                                {dayNum}
                            </div>
                        </button>
                    );
                })}
            </div>
        );
    };

    public render(): React.ReactElement {
        const { onToday, onNavigate, dateRangeText } = this.props;
        const { isDatePickerOpen, searchQuery } = this.state;

        const toolbarTopStyle = mergeStyles({
            display: 'flex',
            alignItems: 'center',
            gap: 12,
            marginBottom: 8,
            justifyContent: 'space-between'
        });

        const toolbarLeftStyle = mergeStyles({
            display: 'flex',
            alignItems: 'center',
            gap: 8
        });

        const toolbarArrowStyle = mergeStyles({
            display: 'flex',
            alignItems: 'center',
            gap: 8
        });

        return (
            <div style={{ marginBottom: 8 }}>
                {/* Top toolbar: Today, Date Picker, and Search */}
                <div className={toolbarTopStyle}>
                    <div className={toolbarLeftStyle}>
                        <DefaultButton
                            text="Today"
                            onClick={onToday}
                            iconProps={{ iconName: 'GotoToday' }}
                        />
                        <div className={toolbarArrowStyle}>
                            <ActionButton
                                iconProps={{ iconName: 'ChevronLeft' }}
                                onClick={() => onNavigate('prev')}
                                ariaLabel={this.getNavigationAriaLabel('prev')}
                                title={this.getNavigationAriaLabel('prev')}
                            /></div>
                        <div ref={this.datePickerButtonRef}>
                            <DefaultButton
                                text={dateRangeText}
                                onClick={() => this.setState({ isDatePickerOpen: !isDatePickerOpen })}
                                iconProps={{ iconName: 'Calendar' }}
                                style={{ minWidth: 220 }}
                            />
                        </div>
                        {isDatePickerOpen && (
                            <Callout
                                target={this.datePickerButtonRef}
                                onDismiss={() => this.setState({ isDatePickerOpen: false })}
                                directionalHint={4} // BottomLeftEdge
                            >
                                {this.getDatePickerContent()}
                            </Callout>
                        )}
                    </div>
                    <div className={toolbarArrowStyle}>
                        <ActionButton
                            iconProps={{ iconName: 'ChevronRight' }}
                            onClick={() => onNavigate('next')}
                            ariaLabel={this.getNavigationAriaLabel('next')}
                            title={this.getNavigationAriaLabel('next')}
                        />
                    </div>
                    <SearchBox
                        placeholder={this.getSearchPlaceholder()}
                        value={searchQuery}
                        onChange={this.handleSearchChange}
                        style={{ width: 250, flexShrink: 0 }}
                    />
                </div>
                
                {/* Quick day selector - only for day view */}
                {this.renderDaySelector()}
            </div>
        );
    }
}
