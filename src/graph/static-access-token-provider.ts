import type { AccessTokenProvider } from "./device-code-token-provider.js";

export class StaticAccessTokenProvider implements AccessTokenProvider {
  constructor(private readonly accessToken: string) {}

  async getAccessToken(): Promise<string> {
    return this.accessToken;
  }
}
