import { InteractiveBrowserCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';

export class OutlookAuthManager {
  constructor(clientId, tenantId) {
    this.clientId = clientId;
    this.tenantId = tenantId;
    this.credential = null;
    this.graphClient = null;
    this.isAuthenticated = false;
  }

  async authenticate() {
    try {
      this.credential = new InteractiveBrowserCredential({
        clientId: this.clientId,
        tenantId: this.tenantId,
        redirectUri: 'http://localhost:8400',
      });

      const authProvider = new TokenCredentialAuthenticationProvider(this.credential, {
        scopes: [
          'Mail.Read',
          'Mail.ReadWrite',
          'Mail.Send',
          'Calendars.Read',
          'Calendars.ReadWrite',
          'Contacts.Read',
          'Contacts.ReadWrite',
          'Tasks.Read',
          'Tasks.ReadWrite',
          'User.Read',
        ],
      });

      this.graphClient = Client.initWithMiddleware({
        authProvider,
      });

      const user = await this.graphClient.api('/me').get();
      this.isAuthenticated = true;
      
      return {
        success: true,
        user: {
          id: user.id,
          displayName: user.displayName,
          mail: user.mail || user.userPrincipalName,
        },
      };
    } catch (error) {
      this.isAuthenticated = false;
      return {
        success: false,
        error: error.message,
      };
    }
  }

  async ensureAuthenticated() {
    if (!this.isAuthenticated) {
      const result = await this.authenticate();
      if (!result.success) {
        throw new Error(`Authentication failed: ${result.error}`);
      }
    }
    return this.graphClient;
  }

  getGraphClient() {
    if (!this.graphClient) {
      throw new Error('Not authenticated. Call authenticate() first.');
    }
    return this.graphClient;
  }
}