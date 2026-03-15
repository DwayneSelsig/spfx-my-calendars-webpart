import { CalendarSourceType } from './ICalendarSettings';

declare module '@pnp/spfx-controls-react/lib/controls/calendar/models/IEvents' {
  interface IEvent {
    // Required by our multi-source calendar
    sourceId: string;
    
    // Source metadata for display
    sourceType?: CalendarSourceType;
    showSourceLogo?: boolean;
    sourceIconName?: string;
    
    // Planner-specific fields
    percentComplete?: number;
    
    // Teams Shifts-specific fields
    isDraft?: boolean;

    // Hex color used by our custom renderers
    colorHex?: string;
  }
}

// This file augments the PnP IEvent interface with our custom properties
// No separate domain model needed - IEvent becomes our canonical event type
// Note: color field already exists in IEvent as AvatarNamedColor but we use hex strings
export {};
