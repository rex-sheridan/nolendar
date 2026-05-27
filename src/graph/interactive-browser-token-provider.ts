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
    let token;

    try {
      token = await this.credential.getToken(this.scopes);
    } catch (error) {
      throw formatInteractiveBrowserError(error);
    }

    if (!token?.token) {
      throw new Error("Microsoft interactive-browser authentication succeeded without returning an access token.");
    }

    return token.token;
  }
}

function formatInteractiveBrowserError(error: unknown): Error {
  const message = error instanceof Error ? error.message : String(error);

  if (message.includes("AADSTS7000218")) {
    return new Error(
      [
        "Microsoft interactive-browser authentication failed because the configured app requires a client secret.",
        "Use `microsoft.authMode: auth_code` with `MICROSOFT_CLIENT_SECRET`, or configure the app registration as a public client for interactive/device-code auth.",
        "",
        `Original error: ${message}`,
      ].join("\n"),
    );
  }

  if (message.includes("AADSTS65002")) {
    return new Error(
      [
        "Microsoft interactive-browser authentication failed because the selected client app cannot request Microsoft Graph tokens in this tenant.",
        "Set `MICROSOFT_CLIENT_ID` to your own app registration with delegated Microsoft Graph permissions such as `Calendars.Read` and `User.Read`.",
        "",
        `Original error: ${message}`,
      ].join("\n"),
    );
  }

  return error instanceof Error ? error : new Error(message);
}
