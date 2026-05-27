import { DeviceCodeCredential } from "@azure/identity";

import type { GraphDeviceCodeAuthConfig } from "./auth.js";

export interface AccessTokenProvider {
  getAccessToken(): Promise<string>;
}

export interface DeviceCodePromptTarget {
  log(message: string): void;
}

interface DeviceCodeInfo {
  message?: string;
  userCode?: string;
  verificationUri?: string;
}

export class DeviceCodeTokenProvider implements AccessTokenProvider {
  private readonly credential: DeviceCodeCredential;
  private readonly scopes: string[];

  constructor(config: GraphDeviceCodeAuthConfig, promptTarget: DeviceCodePromptTarget) {
    this.credential = new DeviceCodeCredential({
      tenantId: config.tenantId,
      clientId: config.clientId,
      userPromptCallback: (info: DeviceCodeInfo) => {
        const message = formatDeviceCodePrompt(info);

        if (message) {
          promptTarget.log(message);
        }
      },
    });
    this.scopes = config.scopes;
  }

  async getAccessToken(): Promise<string> {
    let token;

    try {
      token = await this.credential.getToken(this.scopes);
    } catch (error) {
      throw formatDeviceCodeError(error);
    }

    if (!token?.token) {
      throw new Error("Microsoft Graph authentication succeeded without returning an access token.");
    }

    return token.token;
  }
}

function formatDeviceCodePrompt(info: DeviceCodeInfo): string | null {
  if (typeof info.message === "string" && info.message.trim() !== "") {
    return info.message;
  }

  if (typeof info.userCode === "string" && typeof info.verificationUri === "string") {
    return `Open ${info.verificationUri} and enter code ${info.userCode}`;
  }

  return "Open https://microsoft.com/devicelogin and complete Microsoft sign-in.";
}

function formatDeviceCodeError(error: unknown): Error {
  const message = error instanceof Error ? error.message : String(error);

  if (message.includes("invalid_grant")) {
    return new Error(
      [
        "Microsoft device-code authentication failed with `invalid_grant`.",
        "Common causes:",
        "- the app registration does not have public client flows enabled",
        "- `microsoft.tenant` does not match the type of account you used to sign in",
        "- the device-code sign-in was not completed successfully",
        "- your tenant blocked the flow with Conditional Access",
        "- if your app registration is a web/confidential app, use `microsoft.authMode: auth_code` with `MICROSOFT_CLIENT_SECRET` instead",
        "",
        `Original error: ${message}`,
      ].join("\n"),
    );
  }

  return error instanceof Error ? error : new Error(message);
}
