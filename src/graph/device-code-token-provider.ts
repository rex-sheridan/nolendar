import { DeviceCodeCredential } from "@azure/identity";

import type { GraphAuthConfig } from "./auth.js";

export interface AccessTokenProvider {
  getAccessToken(): Promise<string>;
}

export interface DeviceCodePromptTarget {
  log(message: string): void;
}

interface DeviceCodeInfo {
  message: string;
}

export class DeviceCodeTokenProvider implements AccessTokenProvider {
  private readonly credential: DeviceCodeCredential;
  private readonly scopes: string[];

  constructor(config: GraphAuthConfig, promptTarget: DeviceCodePromptTarget) {
    this.credential = new DeviceCodeCredential({
      tenantId: config.tenantId,
      clientId: config.clientId,
      userPromptCallback: (info: DeviceCodeInfo) => {
        promptTarget.log(info.message);
      },
    });
    this.scopes = config.scopes;
  }

  async getAccessToken(): Promise<string> {
    const token = await this.credential.getToken(this.scopes);

    if (!token?.token) {
      throw new Error("Microsoft Graph authentication succeeded without returning an access token.");
    }

    return token.token;
  }
}
