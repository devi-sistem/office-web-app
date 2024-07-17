import { Client } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";

export class GraphService {
  private client: Client;

  constructor(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
    this.client = Client.initWithMiddleware({
      authProvider: authProvider
    });
  }

  async getMyDocuments() {
    return await this.client.api('/me/drive/root/children').get();
  }

  // Add more methods for different Office operations
}