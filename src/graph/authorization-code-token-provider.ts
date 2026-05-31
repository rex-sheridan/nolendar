import { createServer, type Server } from "node:http";
import { execFile } from "node:child_process";
import os from "node:os";
import { promisify } from "node:util";

import { ConfidentialClientApplication, CryptoProvider, type AccountInfo } from "@azure/msal-node";

import type { GraphAuthorizationCodeAuthConfig } from "./auth.js";
import type { AccessTokenProvider } from "./device-code-token-provider.js";
import { FileTokenCachePlugin } from "./file-token-cache-plugin.js";

const execFileAsync = promisify(execFile);
const TOKEN_EXPIRY_BUFFER_MS = 5 * 60 * 1000;

export interface AuthorizationCodePromptTarget {
  log(message: string): void;
}

interface AuthorizationCodeApp {
  getAuthCodeUrl(args: {
    scopes: string[];
    redirectUri: string;
    codeChallenge: string;
    codeChallengeMethod: "S256";
  }): Promise<string>;
  acquireTokenByCode(args: {
    code: string;
    scopes: string[];
    redirectUri: string;
    codeVerifier?: string;
  }): Promise<{
    accessToken?: string;
    expiresOn?: Date | null;
  } | null>;
  acquireTokenSilent?(args: {
    account: AccountInfo;
    scopes: string[];
  }): Promise<{
    accessToken?: string;
    expiresOn?: Date | null;
  } | null>;
  getTokenCache?(): {
    getAllAccounts(): Promise<AccountInfo[]>;
  };
}

interface CachedToken {
  accessToken: string;
  expiresOnTimestamp: number;
}

export class AuthorizationCodeTokenProvider implements AccessTokenProvider {
  private readonly app: AuthorizationCodeApp;
  private readonly authorize: typeof getAuthorizationCode;
  private cachedToken?: CachedToken;
  private tokenPromise?: Promise<CachedToken>;

  constructor(
    private readonly config: GraphAuthorizationCodeAuthConfig,
    private readonly promptTarget: AuthorizationCodePromptTarget,
    options: {
      app?: AuthorizationCodeApp;
      authorize?: typeof getAuthorizationCode;
      now?: () => number;
    } = {},
  ) {
    this.app =
      options.app ??
      new ConfidentialClientApplication({
        auth: {
          clientId: config.clientId,
          clientSecret: config.clientSecret,
          authority: `https://login.microsoftonline.com/${config.tenantId}`,
        },
        cache: {
          cachePlugin: new FileTokenCachePlugin(config.tokenCachePath),
        },
      });
    this.authorize = options.authorize ?? getAuthorizationCode;
    this.now = options.now ?? Date.now;
  }

  private readonly now: () => number;

  async getAccessToken(): Promise<string> {
    if (this.cachedToken && !isExpiringSoon(this.cachedToken, this.now())) {
      return this.cachedToken.accessToken;
    }

    this.tokenPromise ??= this.acquireAccessToken();

    try {
      this.cachedToken = await this.tokenPromise;
      return this.cachedToken.accessToken;
    } finally {
      this.tokenPromise = undefined;
    }
  }

  private async acquireAccessToken(): Promise<CachedToken> {
    const silentToken = await this.acquireSilentAccessToken();

    if (silentToken) {
      return silentToken;
    }

    const { code, codeVerifier } = await this.authorize(this.app, this.config, this.promptTarget);
    const token = await this.app.acquireTokenByCode({
      code,
      scopes: this.config.scopes,
      redirectUri: this.config.redirectUri,
      codeVerifier,
    });

    if (!token?.accessToken) {
      throw new Error("Microsoft Graph authorization-code authentication succeeded without returning an access token.");
    }

    return {
      accessToken: token.accessToken,
      expiresOnTimestamp: token.expiresOn?.getTime() ?? Number.MAX_SAFE_INTEGER,
    };
  }

  private async acquireSilentAccessToken(): Promise<CachedToken | null> {
    if (!this.app.acquireTokenSilent || !this.app.getTokenCache) {
      return null;
    }

    const accounts = await this.app.getTokenCache().getAllAccounts();
    const account = accounts[0];

    if (!account) {
      return null;
    }

    try {
      const token = await this.app.acquireTokenSilent({
        account,
        scopes: this.config.scopes,
      });

      if (!token?.accessToken) {
        return null;
      }

      return {
        accessToken: token.accessToken,
        expiresOnTimestamp: token.expiresOn?.getTime() ?? Number.MAX_SAFE_INTEGER,
      };
    } catch {
      return null;
    }
  }
}

function isExpiringSoon(token: CachedToken, now: number): boolean {
  return token.expiresOnTimestamp - now <= TOKEN_EXPIRY_BUFFER_MS;
}

async function getAuthorizationCode(
  app: AuthorizationCodeApp,
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
        response.end("Microsoft sign-in failed. You can close this tab.", () => {
          closeCallbackServer(server)
            .then(() => {
              reject(new Error(`Microsoft sign-in failed: ${error}${errorDescription ? ` (${errorDescription})` : ""}`));
            })
            .catch(reject);
        });
        return;
      }

      if (!code) {
        response.statusCode = 400;
        response.end("Missing authorization code. You can close this tab.");
        return;
      }

      response.statusCode = 200;
      response.setHeader("Content-Type", "text/plain; charset=utf-8");
      response.end("Microsoft sign-in completed. You can close this tab and return to nolendar.", () => {
        closeCallbackServer(server)
          .then(() => {
            resolve(code);
          })
          .catch(reject);
      });
    });

    server.once("error", (error) => {
      reject(error);
    });

    server.listen(Number(url.port || "80"), url.hostname);
  });
}

async function closeCallbackServer(server: Server): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    server.close((error) => {
      if (error) {
        reject(error);
        return;
      }

      resolve();
    });
    server.closeIdleConnections();
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
