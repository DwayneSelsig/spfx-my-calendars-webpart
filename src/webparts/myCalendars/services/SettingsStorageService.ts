import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import { ICalendarSettings } from '../models/ICalendarSettings';

/**
 * Service to persist calendar settings to OneDrive's App Folder
 * Uses Files.ReadWrite.AppFolder permission scope
 * https://learn.microsoft.com/en-us/graph/onedrive-sharepoint-appfolder
 */
export class SettingsStorageService {
  private readonly APP_FOLDER_NAME = 'SPFx-My-Calendar-Webpart';
  private readonly SETTINGS_FILE = 'calendar-settings.json';
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

  /**
   * Save settings to App Folder
   */
  public async saveSettings(settings: ICalendarSettings): Promise<boolean> {
    try {
      await this.ensureAppFolderExists();
      const appFolderPath = `/me/drive/special/approot:/${this.APP_FOLDER_NAME}/${this.SETTINGS_FILE}:/content`;
      const client = await this.getGraphClient();

      console.log('Saving settings to App Folder:', appFolderPath);

      await client
        .api(appFolderPath)
        .header('Content-Type', 'application/json')
        .put(settings);

      console.log('Settings saved successfully');
      return true;
    } catch (error) {
      console.error('Error saving settings to App Folder:', error);
      return false;
    }
  }

  /**
   * Load settings from App Folder
   */
  public async loadSettings(): Promise<ICalendarSettings | undefined> {
    try {
      const appFolderPath = `/me/drive/special/approot:/${this.APP_FOLDER_NAME}/${this.SETTINGS_FILE}:/content`;
      const client = await this.getGraphClient();

      console.log('Loading settings from App Folder:', appFolderPath);

      const data = await client
        .api(appFolderPath)
        .get();

      console.log('Settings loaded successfully:', data);
      return data as ICalendarSettings;
    } catch (error) {
      if (this.isNotFoundError(error)) {
        console.log('Settings file not found, using defaults');
        return undefined;
      }
      console.error('Error loading settings from App Folder:', error);
      return undefined;
    }
  }

  /**
   * Delete settings file from App Folder
   */
  public async deleteSettings(): Promise<boolean> {
    try {
      const appFolderPath = `/me/drive/special/approot:/${this.APP_FOLDER_NAME}/${this.SETTINGS_FILE}:`;
      const client = await this.getGraphClient();

      await client.api(appFolderPath).delete();

      console.log('Settings deleted successfully');
      return true;
    } catch (error) {
      if (this.isNotFoundError(error)) {
        console.log('Settings already deleted or not found');
        return true;
      }
      console.error('Error deleting settings:', error);
      return false;
    }
  }
}

