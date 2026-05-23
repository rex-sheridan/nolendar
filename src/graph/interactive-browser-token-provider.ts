import { InteractiveBrowserCredential } from "@azure/identity";

import type { GraphInteractiveBrowserAuthConfig } from "./auth.js";
import type { AccessTokenProvider } from "./device-code-token-provider.js";

export class InteractiveBrowserTokenProvider implements AccessTokenProvider {
  private readonly credential: InteractiveBrowserCredential;
  private readonly scopes: string[];

  constructor(config: GraphInteractiveBrowserAuthConfig) {
    this.credential = new InteractiveBrowserCredential({
      tenantId: config.tenantId,
      clientId: config.clientId,
    });
    this.scopes = config.scopes;
  }

  async getAccessToken(): Promise<string> {
    const token = await this.credential.getToken(this.scopes);

    if (!token?.token) {
      throw new Error("Microsoft interactive-browser authentication succeeded without returning an access token.");
    }

    return token.token;
  }
}
