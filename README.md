# My Calendars SPFx Web Part

## Summary

A fully customizable calendar web part for SharePoint built with SPFx and Fluent UI. This web part allows you to aggregate and display multiple ICS calendar feeds in different views including Day, Week, Month, and Schedule.

## Features

### Multiple Calendar Views
- **Day View**: Shows appointments for a single day with hourly time slots
- **Week View**: Displays a full week with configurable weekend visibility
- **Month View**: Traditional month calendar with appointment indicators
- **Schedule View**: List-based view showing appointments chronologically

### Calendar Source Management
- Add multiple ICS calendar feeds via URL
- Configure custom colors for each calendar source
- Enable/disable individual calendar sources
- Support for standard ICS/iCalendar format

### Customization Options
- Configurable work hours (start/end time)
- Adjustable time slot duration (15-60 minutes)
- Show/hide weekends
- First day of week setting
- Default view selection
- Color customization using CSS color-mix for semi-transparent backgrounds

### User Experience
- Smooth animations and transitions
- Skeleton loading states (Shimmer effects)
- Responsive design
- CommandBar for easy navigation
- Debounced settings changes (500ms)
- Regional date/time formatting support

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.22.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

None

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| spfx-my-calendars-webpart | Your Name |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | January 7, 2026 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install`
  - `npm run start` (for local testing)
  - `npm run build` (for production build)

## Configuration

### Adding Calendar Sources

1. Edit the web part in SharePoint
2. Open the property pane
3. Under "Calendar Sources", click "Add Calendar Source"
4. Configure:
   - **Name**: Display name for the calendar
   - **ICS URL**: Full URL to the ICS calendar feed
   - **Color**: Hex color code for this calendar's events
   - **Enabled**: Toggle to show/hide this calendar

### View Settings

Configure the following in the property pane:

- **Default View**: Choose Day, Week, Month, or Schedule
- **Show Weekends**: Toggle weekend display in Week view
- **Start Hour**: First hour shown in Day/Week views (0-23)
- **End Hour**: Last hour shown in Day/Week views (0-23)
- **Slot Duration**: Time slot size in minutes (15, 30, 45, or 60)
- **First Day of Week**: 0=Sunday, 1=Monday, etc.

### ICS Proxy (CORS)

Sommige ICS-bestanden kunnen niet direct opgehaald worden door CORS-beperkingen. De webpart probeert altijd eerst de rechtstreekse URL. Als dat faalt (bijv. CORS), worden proxies geprobeerd in de volgorde die de beheerder instelt.

In het property pane kun je:
- Een of meer proxy-opties inschakelen:
  - Eigen proxy (met invulveld voor de proxy-URL)
  - Openbare proxy: `https://www.whateverorigin.org/`
- De fallback-volgorde bepalen (Eerste/Tweede proxy)

Tip: Voor het opzetten van een eigen CORS-proxy kun je ook deze oplossing bekijken: https://github.com/Zibri/cloudflare-cors-anywhere

## Technical Details

### Architecture

```
src/webparts/myCalendars/
├── components/
│   ├── views/
│   │   ├── DayView.tsx
│   │   ├── WeekView.tsx
│   │   ├── MonthView.tsx
│   │   ├── ScheduleView.tsx
│   │   └── CalendarView.module.scss
│   ├── MyCalendars.tsx (Main component)
│   ├── MyCalendars.module.scss
│   └── IMyCalendarsProps.ts
├── models/
│   ├── IAppointment.ts (Calendar event model)
│   └── ICalendarSettings.ts (Settings interface)
├── services/
│   └── IcsParser.ts (ICS file parser)
├── propertyPane/
│   └── PropertyPaneCalendarSources.tsx (Custom property pane control)
└── MyCalendarsWebPart.ts
```

### Settings Storage

Settings are stored as a JSON string in the web part properties:

```json
{
  "defaultView": "month",
  "sources": [
    {
      "id": "source_123",
      "name": "My Calendar",
      "url": "https://example.com/calendar.ics",
      "color": "#0078d4",
      "isEnabled": true
    }
  ],
  "showWeekends": true,
  "startHour": 8,
  "endHour": 18,
  "slotDuration": 30,
  "firstDayOfWeek": 1
}
```

### ICS Parser

The custom ICS parser supports:
- Standard VEVENT format
- All-day events (VALUE=DATE)
- UTC and local timezones
- Event properties: UID, SUMMARY, DESCRIPTION, LOCATION, DTSTART, DTEND, ORGANIZER
- Line continuation
- Escaped values

### Styling

Uses Fluent UI components and CSS variables for theming:
- Respects SharePoint theme colors
- CSS `color-mix()` for semi-transparent event backgrounds
- Smooth animations with CSS transitions
- Responsive layouts with Flexbox and CSS Grid

## Features

This web part illustrates the following concepts:

- Custom ICS/iCalendar feed parsing
- Multiple calendar aggregation
- Fluent UI integration
- Custom property pane controls
- SharePoint theme integration
- Responsive calendar layouts
- Debounced user input
- Skeleton loading states
- Language-independent date/time formatting

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)
- [Heft Documentation](https://heft.rushstack.io/)
- [Fluent UI React](https://developer.microsoft.com/fluentui)
- [iCalendar (RFC 5545)](https://datatracker.ietf.org/doc/html/rfc5545)