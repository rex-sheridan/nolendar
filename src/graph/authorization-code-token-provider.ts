import { createServer } from "node:http";
import { execFile } from "node:child_process";
import os from "node:os";
import { promisify } from "node:util";

import { ConfidentialClientApplication, CryptoProvider } from "@azure/msal-node";

import type { GraphAuthorizationCodeAuthConfig } from "./auth.js";
import type { AccessTokenProvider } from "./device-code-token-provider.js";

const execFileAsync = promisify(execFile);

export interface AuthorizationCodePromptTarget {
  log(message: string): void;
}

export class AuthorizationCodeTokenProvider implements AccessTokenProvider {
  private readonly app: ConfidentialClientApplication;

  constructor(
    private readonly config: GraphAuthorizationCodeAuthConfig,
    private readonly promptTarget: AuthorizationCodePromptTarget,
  ) {
    this.app = new ConfidentialClientApplication({
      auth: {
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    });
  }

  async getAccessToken(): Promise<string> {
    const { code, codeVerifier } = await getAuthorizationCode(this.app, this.config, this.promptTarget);
    const token = await this.app.acquireTokenByCode({
      code,
      scopes: this.config.scopes,
      redirectUri: this.config.redirectUri,
      codeVerifier,
    });

    if (!token?.accessToken) {
      throw new Error("Microsoft Graph authorization-code authentication succeeded without returning an access token.");
    }

    return token.accessToken;
  }
}

async function getAuthorizationCode(
  app: ConfidentialClientApplication,
  config: GraphAuthorizationCodeAuthConfig,
  promptTarget: AuthorizationCodePromptTarget,
): Promise<{ code: string; codeVerifier?: string }> {
  const pkceCodes = await new CryptoProvider().generatePkceCodes();
  const authUrl = await app.getAuthCodeUrl({
    scopes: config.scopes,
    redirectUri: config.redirectUri,
    codeChallenge: pkceCodes.challenge,
    codeChallengeMethod: "S256",
  });

  const callback = waitForAuthorizationCode(config.redirectUri);
  promptTarget.log(`Opening browser for Microsoft sign-in: ${config.redirectUri}`);
  await openUrl(authUrl, promptTarget);

  const code = await callback;
  return {
    code,
    codeVerifier: pkceCodes.verifier,
  };
}

async function waitForAuthorizationCode(redirectUri: string): Promise<string> {
  const url = new URL(redirectUri);

  if (url.protocol !== "http:" || url.hostname !== "localhost") {
    throw new Error("Authorization-code auth currently requires `MICROSOFT_REDIRECT_URI` to use http://localhost.");
  }

  return new Promise<string>((resolve, reject) => {
    const server = createServer((request, response) => {
      if (!request.url) {
        response.statusCode = 400;
        response.end("Missing callback URL.");
        return;
      }

      const callbackUrl = new URL(request.url, redirectUri);
      const code = callbackUrl.searchParams.get("code");
      const error = callbackUrl.searchParams.get("error");
      const errorDescription = callbackUrl.searchParams.get("error_description");

      if (error) {
        response.statusCode = 400;
        response.end("Microsoft sign-in failed. You can close this tab.");
        server.close();
        reject(new Error(`Microsoft sign-in failed: ${error}${errorDescription ? ` (${errorDescription})` : ""}`));
        return;
      }

      if (!code) {
        response.statusCode = 400;
        response.end("Missing authorization code. You can close this tab.");
        return;
      }

      response.statusCode = 200;
      response.setHeader("Content-Type", "text/plain; charset=utf-8");
      response.end("Microsoft sign-in completed. You can close this tab and return to nolendar.");
      server.close();
      resolve(code);
    });

    server.once("error", (error) => {
      reject(error);
    });

    server.listen(Number(url.port || "80"), url.hostname);
  });
}

async function openUrl(url: string, promptTarget: AuthorizationCodePromptTarget): Promise<void> {
  const platform = os.platform();

  try {
    if (platform === "darwin") {
      await execFileAsync("open", [url]);
      return;
    }

    if (platform === "win32") {
      await execFileAsync("rundll32", ["url.dll,FileProtocolHandler", url]);
      return;
    }

    await execFileAsync("xdg-open", [url]);
  } catch {
    promptTarget.log(`Open this URL in your browser to continue Microsoft sign-in:\n${url}`);
  }
}
