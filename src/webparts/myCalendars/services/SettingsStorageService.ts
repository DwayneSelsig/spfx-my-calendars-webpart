import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import type { ILegacyCalendarSettings, IUserCalendarSettings } from '../models/ICalendarSettings';
import { normalizeUserCalendarSettings } from './CalendarSettingsService';

/**
 * Stores user settings in the OneDrive App Folder.
 * Uses Files.ReadWrite.AppFolder permission scope.
 */
export class SettingsStorageService {
  private readonly APP_FOLDER_NAME = 'SPFx-My-Calendar-Webpart';
  private readonly USER_SETTINGS_FILE = 'user-calendar-settings.json';
  private readonly LEGACY_SETTINGS_FILE = 'calendar-settings.json';
  private graphClientFactory: MSGraphClientFactory;

  constructor(graphClientFactory: MSGraphClientFactory) {
    this.graphClientFactory = graphClientFactory;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    return this.graphClientFactory.getClient('3');
  }

  private isNotFoundError(error: unknown): boolean {
    const err = error as { statusCode?: number; code?: string } | undefined;
    return err?.statusCode === 404 || err?.code === 'itemNotFound';
  }

  private isConflictError(error: unknown): boolean {
    const err = error as { statusCode?: number; code?: string } | undefined;
    return err?.statusCode === 409 || err?.code === 'nameAlreadyExists';
  }

  private async ensureAppFolderExists(): Promise<void> {
    const client = await this.getGraphClient();
    try {
      await client.api('/me/drive/special/approot/children').post({
        name: this.APP_FOLDER_NAME,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail'
      });
    } catch (error) {
      if (this.isConflictError(error)) {
        return;
      }
      throw error;
    }
  }

  private async readJsonFile<T>(fileName: string): Promise<T | undefined> {
    try {
      const appFolderPath = `/me/drive/special/approot:/${this.APP_FOLDER_NAME}/${fileName}:/content`;
      const client = await this.getGraphClient();
      const data = await client.api(appFolderPath).get();
      return data as T;
    } catch (error) {
      if (this.isNotFoundError(error)) {
        return undefined;
      }
      console.error(`Error loading ${fileName} from App Folder:`, error);
      return undefined;
    }
  }

  private async writeJsonFile(fileName: string, data: unknown): Promise<boolean> {
    try {
      await this.ensureAppFolderExists();
      const appFolderPath = `/me/drive/special/approot:/${this.APP_FOLDER_NAME}/${fileName}:/content`;
      const client = await this.getGraphClient();

      await client
        .api(appFolderPath)
        .header('Content-Type', 'application/json')
        .put(data);

      return true;
    } catch (error) {
      console.error(`Error saving ${fileName} to App Folder:`, error);
      return false;
    }
  }

  private async deleteJsonFile(fileName: string): Promise<boolean> {
    try {
      const appFolderPath = `/me/drive/special/approot:/${this.APP_FOLDER_NAME}/${fileName}:`;
      const client = await this.getGraphClient();
      await client.api(appFolderPath).delete();
      return true;
    } catch (error) {
      if (this.isNotFoundError(error)) {
        return true;
      }
      console.error(`Error deleting ${fileName} from App Folder:`, error);
      return false;
    }
  }

  public async saveUserSettings(settings: IUserCalendarSettings): Promise<boolean> {
    return this.writeJsonFile(this.USER_SETTINGS_FILE, settings);
  }

  public async loadUserSettings(): Promise<IUserCalendarSettings | undefined> {
    const data = await this.readJsonFile<unknown>(this.USER_SETTINGS_FILE);
    return normalizeUserCalendarSettings(data);
  }

  public async loadLegacySettings(): Promise<ILegacyCalendarSettings | undefined> {
    return this.readJsonFile<ILegacyCalendarSettings>(this.LEGACY_SETTINGS_FILE);
  }

  public async deleteUserSettings(): Promise<boolean> {
    const [deletedCurrent, deletedLegacy] = await Promise.all([
      this.deleteJsonFile(this.USER_SETTINGS_FILE),
      this.deleteJsonFile(this.LEGACY_SETTINGS_FILE)
    ]);

    return deletedCurrent && deletedLegacy;
  }
}
