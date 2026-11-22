import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphRequestScheduler, type GraphRequestScope } from './GraphRequestScheduler';

type SdkRequest = ReturnType<MSGraphClientV3['api']>;

export interface GraphRequest {
  header(name: string, value: string): GraphRequest;
  query(values: Parameters<SdkRequest['query']>[0]): GraphRequest;
  select(fields: string): GraphRequest;
  top(count: number): GraphRequest;
  expand(fields: string): GraphRequest;
  get(): ReturnType<SdkRequest['get']>;
  post(body: unknown): ReturnType<SdkRequest['post']>;
  put(body: unknown): ReturnType<SdkRequest['put']>;
  patch(body: unknown): ReturnType<SdkRequest['patch']>;
  delete(): ReturnType<SdkRequest['delete']>;
}

export interface GraphClient { api(path: string): GraphRequest }

class ScheduledGraphRequest implements GraphRequest {
  private readonly configure: Array<(request: GraphRequest) => GraphRequest> = [];

  public constructor(private readonly execute: <T>(operation: (request: GraphRequest) => Promise<T>) => Promise<T>) {}

  public header(name: string, value: string): GraphRequest { this.configure.push(request => request.header(name, value)); return this; }
  public query(values: Parameters<SdkRequest['query']>[0]): GraphRequest { this.configure.push(request => request.query(values)); return this; }
  public select(fields: string): GraphRequest { this.configure.push(request => request.select(fields)); return this; }
  public top(count: number): GraphRequest { this.configure.push(request => request.top(count)); return this; }
  public expand(fields: string): GraphRequest { this.configure.push(request => request.expand(fields)); return this; }
  public get(): ReturnType<SdkRequest['get']> { return this.send(request => request.get()); }
  public post(body: unknown): ReturnType<SdkRequest['post']> { return this.send(request => request.post(body)); }
  public put(body: unknown): ReturnType<SdkRequest['put']> { return this.send(request => request.put(body)); }
  public patch(body: unknown): ReturnType<SdkRequest['patch']> { return this.send(request => request.patch(body)); }
  public delete(): ReturnType<SdkRequest['delete']> { return this.send(request => request.delete()); }

  private send<T>(operation: (request: GraphRequest) => Promise<T>): Promise<T> {
    return this.execute(request => operation(this.configure.reduce((current, configure) => configure(current), request)));
  }
}

export class GraphRequestExecutor implements GraphClient {
  private client?: Promise<GraphClient>;

  public constructor(private readonly connect: () => Promise<GraphClient>, private readonly scheduler: GraphRequestScheduler) {}

  public api(path: string): GraphRequest { return this.scoped().api(path); }

  public scoped(scope?: GraphRequestScope): GraphClient {
    return { api: path => new ScheduledGraphRequest(operation => this.execute(path, operation, scope)) };
  }

  private async execute<T>(path: string, operation: (request: GraphRequest) => Promise<T>, scope?: GraphRequestScope): Promise<T> {
    this.client = this.client || this.connect().catch(error => {
      this.client = undefined;
      throw error;
    });
    const client = await this.client;
    return this.scheduler.run(() => operation(client.api(path)), scope);
  }
}

interface GraphPage<T> { value?: T[]; '@odata.nextLink'?: string }

export async function forEachGraphPage<T>(
  client: GraphClient,
  path: string,
  onPage: (items: T[]) => void | Promise<void>,
  configure?: (request: GraphRequest) => GraphRequest,
  headers: Record<string, string> = {}
): Promise<void> {
  let next: string | undefined = path;
  const seen = new Set<string>();
  while (next) {
    if (seen.has(next)) throw new Error('Graph returned a repeated continuation link.');
    let request = client.api(next);
    if (!seen.size && configure) request = configure(request);
    for (const name of Object.keys(headers)) request = request.header(name, headers[name]);
    seen.add(next);
    const page: GraphPage<T> = await request.get();
    await onPage(page.value || []);
    next = page['@odata.nextLink'];
  }
}

export async function collectGraphPages<T>(
  client: GraphClient, path: string, configure?: (request: GraphRequest) => GraphRequest, headers?: Record<string, string>
): Promise<T[]> {
  const items: T[] = [];
  await forEachGraphPage<T>(client, path, page => { items.push(...page); }, configure, headers);
  return items;
}
