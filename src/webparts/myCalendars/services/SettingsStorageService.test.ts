jest.mock('@microsoft/sp-http', () => ({}));

import { SettingsStorageService } from './SettingsStorageService';
import type { MSGraphClientFactory } from '@microsoft/sp-http';

describe('personal reset storage boundary', () => {
  it('deletes legacy before current', async () => {
    const paths: string[] = [];
    const client = { api: (path: string) => ({ delete: async () => { paths.push(path); } }) };
    const factory = { getClient: async () => client } as unknown as MSGraphClientFactory;
    expect(await new SettingsStorageService(factory).deleteUserSettings()).toBe(true);
    expect(paths[0]).toContain('/calendar-settings.json:');
    expect(paths[1]).toContain('/user-calendar-settings.json:');
  });

  it('keeps current when legacy cleanup fails and treats missing files as success', async () => {
    const failure = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    const remove = jest.fn().mockRejectedValueOnce({ statusCode: 403 });
    const api = jest.fn(() => ({ delete: remove }));
    const factory = { getClient: async () => ({ api }) } as unknown as MSGraphClientFactory;
    const storage = new SettingsStorageService(factory);
    expect(await storage.deleteUserSettings()).toBe(false);
    expect(api).toHaveBeenCalledTimes(1);
    remove.mockRejectedValue({ statusCode: 404 });
    expect(await storage.deleteUserSettings()).toBe(true);
    failure.mockRestore();
  });
});
