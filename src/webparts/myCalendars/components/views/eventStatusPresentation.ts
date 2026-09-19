import type { CSSProperties } from 'react';
import type { CalendarAvailabilityStatus, CalendarResponseStatus, ICalendarEvent } from '../../models/ICalendarEvent';

export interface IEventStatusPresentation {
  availabilityIconName?: string;
  responseIconName?: string;
  cardStyle: CSSProperties;
  titleStyle: CSSProperties;
}

const availabilityIcons: Record<CalendarAvailabilityStatus, string> = {
  free: 'CircleRing',
  tentative: 'Clock',
  busy: 'CircleFill',
  oof: 'Airplane',
  workingElsewhere: 'Home',
  unknown: 'Unknown'
};

const responseIcons: Record<CalendarResponseStatus, string> = {
  none: 'StatusCircleRing',
  organizer: 'PartyLeader',
  tentativelyAccepted: 'Clock',
  accepted: 'StatusCircleCheckmark',
  declined: 'StatusCircleErrorX',
  notResponded: 'StatusCircleQuestionMark'
};

export function getEventStatusPresentation(event: ICalendarEvent, color: string, baseTintPercent: number = 18): IEventStatusPresentation {
  const showAs = event.showAs;
  const tint = showAs === 'unknown' ? 10 : showAs === 'tentative' || showAs === 'oof' || showAs === 'workingElsewhere' ? 14 : baseTintPercent;
  const outlineStyle = showAs === 'tentative' ? 'dashed' : showAs === 'workingElsewhere' ? 'dotted' : 'solid';
  const outlineColor = 'var(--neutralTertiaryAlt, #c8c6c4)';
  const declined = event.responseStatus === 'declined';
  const normalBackground = `color-mix(in srgb, ${color} ${declined ? Math.max(4, Math.round(tint / 2)) : tint}%, var(--white, #fff))`;
  const diagonalBackground = `repeating-linear-gradient(135deg, color-mix(in srgb, ${color} ${declined ? 9 : 18}%, var(--white, #fff)) 0 5px, color-mix(in srgb, ${color} ${declined ? 4 : 7}%, var(--white, #fff)) 5px 10px)`;
  const crossHatchLine = `color-mix(in srgb, ${color} ${declined ? 14 : 25}%, transparent)`;
  const crossHatchBase = `color-mix(in srgb, ${color} ${declined ? 4 : 7}%, var(--white, #fff))`;
  const crossHatchBackground = `repeating-linear-gradient(45deg, transparent 0 4px, ${crossHatchLine} 4px 5px, transparent 5px 9px), repeating-linear-gradient(-45deg, transparent 0 4px, ${crossHatchLine} 4px 5px, transparent 5px 9px), ${crossHatchBase}`;
  const background = showAs === 'free'
    ? 'transparent'
    : showAs === 'tentative'
      ? diagonalBackground
      : showAs === 'oof'
        ? crossHatchBackground
        : normalBackground;

  return {
    availabilityIconName: showAs ? availabilityIcons[showAs] : undefined,
    responseIconName: event.responseStatus ? responseIcons[event.responseStatus] : undefined,
    cardStyle: {
      background,
      borderTop: showAs ? `1px ${outlineStyle} ${outlineColor}` : undefined,
      borderRight: showAs ? `1px ${outlineStyle} ${outlineColor}` : undefined,
      borderBottom: showAs ? `1px ${outlineStyle} ${outlineColor}` : undefined,
      color: declined ? 'var(--neutralSecondary, #605e5c)' : 'var(--neutralPrimary, #323130)'
    },
    titleStyle: { textDecoration: declined ? 'line-through' : 'none' }
  };
}
