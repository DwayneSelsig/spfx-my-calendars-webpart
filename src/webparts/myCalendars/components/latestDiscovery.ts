/** Keeps discovery callbacks scoped to the most recent request/active dialog. */
export class LatestDiscovery {
  private generation = 0;

  public invalidate(): void {
    this.generation++;
  }

  public async run<T>(
    load: () => Promise<T>,
    callbacks: { start: () => void; success: (value: T) => void; error: (error: unknown) => void; finish: () => void }
  ): Promise<void> {
    const generation = ++this.generation;
    callbacks.start();
    try {
      const value = await load();
      if (generation === this.generation) callbacks.success(value);
    } catch (error) {
      if (generation === this.generation) callbacks.error(error);
    } finally {
      if (generation === this.generation) callbacks.finish();
    }
  }
}
