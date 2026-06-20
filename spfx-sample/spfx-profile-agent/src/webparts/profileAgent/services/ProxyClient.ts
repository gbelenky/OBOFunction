import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';

// The proxy API app registration (api://<proxy-app-client-id>) and its public base URL.
// The browser talks ONLY to the proxy; it fronts the /api/agent/chat endpoint, which
// delegates all profile data retrieval to the Foundry agent (per-user MCP tool).
export const PROXY_RESOURCE = 'api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a';
export const PROXY_BASE = 'https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net';

export interface ChatReply {
  reply: string;
  responseId: string;
  status: string;
  consentUrl?: string;
}

/** Thin client over the OBO proxy's agent-chat endpoint. */
export class ProxyClient {
  private client: AadHttpClient | undefined;
  private previousResponseId: string | undefined = undefined;

  public constructor(private readonly factory: AadHttpClientFactory) {}

  private async getClient(): Promise<AadHttpClient> {
    if (!this.client) {
      this.client = await this.factory.getClient(PROXY_RESOURCE);
    }
    return this.client;
  }

  /**
   * Sends a chat turn to the agent via the proxy. If the agent needs OAuth consent, returns a
   * ChatReply with status 'consent_required' and a consentUrl; open it, then call again unchanged.
   */
  public async ask(message: string): Promise<ChatReply> {
    const client = await this.getClient();
    const res: HttpClientResponse = await client.post(
      `${PROXY_BASE}/api/agent/chat`,
      AadHttpClient.configurations.v1,
      {
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ message, previousResponseId: this.previousResponseId })
      }
    );
    if (!res.ok) {
      throw new Error(`agent ${res.status}: ${await res.text()}`);
    }
    const data = (await res.json()) as ChatReply;
    // Keep the server-side thread for multi-turn state.
    if (data.responseId) {
      this.previousResponseId = data.responseId;
    }
    return data;
  }
}
