jest.mock('@microsoft/sp-http', () => ({}));

import { SettingsStorageService } from './SettingsStorageService';
import type { GraphClient } from './GraphRequestExecutor';

describe('personal reset storage boundary', () => {
  it('deletes legacy before current', async () => {
    const paths: string[] = [];
    const client = { api: (path: string) => ({ delete: async () => { paths.push(path); } }) };
    expect(await new SettingsStorageService(client as unknown as GraphClient).deleteUserSettings()).toBe(true);
    expect(paths[0]).toContain('/calendar-settings.json:');
    expect(paths[1]).toContain('/user-calendar-settings.json:');
  });

  it('keeps current when legacy cleanup fails and treats missing files as success', async () => {
    const failure = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    const remove = jest.fn().mockRejectedValueOnce({ statusCode: 403 });
    const api = jest.fn(() => ({ delete: remove }));
    const storage = new SettingsStorageService({ api } as unknown as GraphClient);
    expect(await storage.deleteUserSettings()).toBe(false);
    expect(api).toHaveBeenCalledTimes(1);
    remove.mockRejectedValue({ statusCode: 404 });
    expect(await storage.deleteUserSettings()).toBe(true);
    failure.mockRestore();
  });
});
