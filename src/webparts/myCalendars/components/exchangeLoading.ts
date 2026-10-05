import type { ICalendarEvent } from '../models/ICalendarEvent';

export interface IExchangeLoadTask {
  sourceId: string;
  load: () => Promise<ICalendarEvent[]>;
}

/** Automatic discovery cannot prevent configured sources from being attempted. */
export async function loadExchangeSources(
  discover: () => Promise<IExchangeLoadTask[]>,
  configured: IExchangeLoadTask[],
  onLoaded: (sourceId: string, events: ICalendarEvent[]) => void,
  onError: (error: unknown, sourceId?: string) => void
): Promise<{ events: ICalendarEvent[]; hadError: boolean }> {
  let hadError = false;
  const load = async (task: IExchangeLoadTask): Promise<ICalendarEvent[]> => {
    try {
      const events = await task.load();
      onLoaded(task.sourceId, events);
      return events;
    } catch (error) {
      hadError = true;
      onError(error, task.sourceId);
      return [];
    }
  };
  const automatic = async (): Promise<ICalendarEvent[][]> => {
    try {
      return await Promise.all((await discover()).map(load));
    } catch (error) {
      hadError = true;
      onError(error);
      return [];
    }
  };
  const [automaticEvents, configuredEvents] = await Promise.all([automatic(), Promise.all(configured.map(load))]);
  return { events: automaticEvents.concat(configuredEvents).reduce<ICalendarEvent[]>((all, events) => all.concat(events), []), hadError };
}
