interface ICachedUserInfo {
  email: string;
  timestamp: number;
}

interface ICachedMailboxSettings {
  mailboxSettings: IMailboxSettings;
  timestamp: number;
}

interface IMeResponse {
  mail?: string;
  userPrincipalName?: string;
}

interface IGetScheduleResponse {
  value?: Array<{
    workingHours?: IWorkingHours;
  }>;
}

export interface IWorkingHours {
  daysOfWeek?: string[];
  startTime?: string;
  endTime?: string;
  timeZone?: {
    name?: string;
  };
}

export interface IMailboxSettings {
  timeZone?: string;
  timeFormat?: string;
  dateFormat?: string;
  workingHours?: IWorkingHours;
}

interface IGraphRequestBuilder {
  select(fields: string): { get(): Promise<IMeResponse | IMailboxSettings> };
  post(body: unknown): Promise<IGetScheduleResponse>;
}

interface IGraphClientLike {
  api(path: string): IGraphRequestBuilder;
}

export class UserHelper {
  private static readonly EMAIL_CACHE_KEY = 'currentUserEmailCache';
  private static readonly MAILBOX_SETTINGS_CACHE_KEY = 'currentUserMailboxSettingsCache';
  private static readonly CACHE_DURATION_MS = 12 * 60 * 60 * 1000; // 12 hours
  
  /**
   * Get current user's email with 12-hour local storage caching
   */
  public static async getCurrentUserEmail(graphClient: IGraphClientLike): Promise<string> {
    try {
      // Check local storage cache first
      const cachedData = this.getCachedUserEmail();
      if (cachedData) {
        return cachedData;
      }
      
      // Fetch from Graph API
      const response = await graphClient
        .api('/me')
        .select('id,userPrincipalName,mail')
        .get() as IMeResponse;
      
      const email = response.mail || response.userPrincipalName || '';
      
      // Cache the result
      if (email) {
        console.log('Fetched user id + principal name + mail from Graph API:', { email });
        this.setCachedUserEmail(email);
      }
      else {
        console.warn('Graph API did not return mail or userPrincipalName for /me endpoint');
      }
      
      return email;
    } catch (error) {
      console.error('Failed to fetch current user:', error);
      return '';
    }
  }

  /**
   * Get current user's mailbox settings with 12-hour local storage caching.
   * Uses /me/calendar/getSchedule with a fixed historical date range to read workingHours only.
   * Requires Microsoft Graph permission: Calendars.ReadBasic (or higher).
   */
  public static async getCurrentUserMailboxSettings(graphClient: IGraphClientLike): Promise<IMailboxSettings | undefined> {
    try {
      const cachedSettings = this.getCachedMailboxSettings();
      if (cachedSettings) {
        return cachedSettings;
      }

      const currentUserEmail = await this.getCurrentUserEmail(graphClient);
      if (!currentUserEmail) {
        console.warn('Unable to resolve current user email for getSchedule request');
        return undefined;
      }

      const response = await graphClient
        .api('/me/calendar/getSchedule')
        .post({
          schedules: [currentUserEmail],
          startTime: {
            dateTime: '1900-03-15T09:00:00',
            timeZone: 'UTC'
          },
          endTime: {
            dateTime: '1900-03-15T18:00:00',
            timeZone: 'UTC'
          }
        });

      const workingHours = response?.value?.[0]?.workingHours;
      const mailboxSettings: IMailboxSettings = {
        workingHours,
        timeZone: workingHours?.timeZone?.name
      };

      if (mailboxSettings.workingHours || mailboxSettings.timeZone) {
        this.setCachedMailboxSettings(mailboxSettings);
      }
      console.log('Fetched mailbox settings from Graph API:', { mailboxSettings });
      return mailboxSettings;
    } catch (error) {
      console.error('Failed to fetch current user mailbox settings:', error);
      return undefined;
    }
  }
  
  /**
   * Check if a given email matches the current user email (case-insensitive)
   */
  public static isEventOrganizer(organizerEmail: string | undefined, currentUserEmail: string): boolean {
    if (!organizerEmail || !currentUserEmail) {
      return false;
    }
    return organizerEmail.toLowerCase() === currentUserEmail.toLowerCase();
  }
  
  /**
   * Get cached user email from local storage if still valid
   */
  private static getCachedUserEmail(): string | null {
    try {
      const cachedJson = localStorage.getItem(this.EMAIL_CACHE_KEY);
      if (!cachedJson) {
        return null;
      }
      
      const cached: ICachedUserInfo = JSON.parse(cachedJson);
      const now = Date.now();
      
      // Check if cache is still valid (within 12 hours)
      if (now - cached.timestamp < this.CACHE_DURATION_MS) {
        return cached.email;
      }
      
      // Cache expired, remove it
      localStorage.removeItem(this.EMAIL_CACHE_KEY);
      return null;
    } catch (error) {
      console.error('Error reading user cache:', error);
      return null;
    }
  }
  
  /**
   * Store user email in local storage with timestamp
   */
  private static setCachedUserEmail(email: string): void {
    try {
      const cacheData: ICachedUserInfo = {
        email,
        timestamp: Date.now()
      };
      localStorage.setItem(this.EMAIL_CACHE_KEY, JSON.stringify(cacheData));
    } catch (error) {
      console.error('Error caching user email:', error);
    }
  }

  /**
   * Get cached mailbox settings from local storage if still valid
   */
  private static getCachedMailboxSettings(): IMailboxSettings | undefined {
    try {
      const cachedJson = localStorage.getItem(this.MAILBOX_SETTINGS_CACHE_KEY);
      if (!cachedJson) {
        return undefined;
      }

      const cached: ICachedMailboxSettings = JSON.parse(cachedJson);
      const now = Date.now();

      if (now - cached.timestamp < this.CACHE_DURATION_MS) {
        return cached.mailboxSettings;
      }

      localStorage.removeItem(this.MAILBOX_SETTINGS_CACHE_KEY);
      return undefined;
    } catch (error) {
      console.error('Error reading mailbox settings cache:', error);
      return undefined;
    }
  }

  /**
   * Store mailbox settings in local storage with timestamp
   */
  private static setCachedMailboxSettings(mailboxSettings: IMailboxSettings): void {
    try {
      const cacheData: ICachedMailboxSettings = {
        mailboxSettings,
        timestamp: Date.now()
      };
      localStorage.setItem(this.MAILBOX_SETTINGS_CACHE_KEY, JSON.stringify(cacheData));
    } catch (error) {
      console.error('Error caching mailbox settings:', error);
    }
  }
  
  /**
   * Clear cached user email (useful for logout or testing)
   */
  public static clearCache(): void {
    try {
      localStorage.removeItem(this.EMAIL_CACHE_KEY);
      localStorage.removeItem(this.MAILBOX_SETTINGS_CACHE_KEY);
    } catch (error) {
      console.error('Error clearing user cache:', error);
    }
  }
}
