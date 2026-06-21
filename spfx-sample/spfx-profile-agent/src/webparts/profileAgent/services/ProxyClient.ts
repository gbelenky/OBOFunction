import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';

// The proxy API app registration (api://<proxy-app-client-id>) and its public base URL.
// The browser talks ONLY to the proxy. The proxy authenticates the user, resolves their
// SharePoint/Graph profile via OBO, injects it as context, and calls the Foundry hosted agent
// on the user's behalf — the front-end stays agnostic of profile and tool logic.
export const PROXY_RESOURCE = 'api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a';
export const PROXY_BASE = 'https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net';

export interface ChatReply {
  reply: string;
  responseId: string;
  status: string;
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
   * Sends a chat turn to the agent via the proxy and returns the agent's reply. Multi-turn
   * state is carried server-side via previousResponseId.
   */
  public async ask(message: string): Promise<ChatReply> {
    return this.post({ message, previousResponseId: this.previousResponseId });
  }

  /**
   * Requests the opening greeting WITHOUT prescribing any wording. The client stays agnostic:
   * it only signals that the chat was opened (greeting:true). The proxy supplies the greeting
   * trigger and the agent owns the greeting text (one short sentence, by first name, no profile dump).
   */
  public async greet(): Promise<ChatReply> {
    return this.post({ message: '', greeting: true });
  }

  private async post(payload: object): Promise<ChatReply> {
    const client = await this.getClient();
    const res: HttpClientResponse = await client.post(
      `${PROXY_BASE}/api/agent/chat`,
      AadHttpClient.configurations.v1,
      {
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(payload)
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
