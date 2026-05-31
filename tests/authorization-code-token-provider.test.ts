import { describe, expect, it, vi } from "vitest";

import { AuthorizationCodeTokenProvider } from "../src/graph/authorization-code-token-provider.js";

const CONFIG = {
  mode: "auth_code" as const,
  tenantId: "organizations",
  clientId: "client-id",
  clientSecret: "client-secret",
  redirectUri: "http://localhost:8787/auth/callback",
  scopes: ["Calendars.Read", "User.Read"],
};

describe("AuthorizationCodeTokenProvider", () => {
  it("reuses a cached access token until it is close to expiry", async () => {
    let now = Date.parse("2026-05-27T20:00:00.000Z");
    const authorize = vi.fn(async () => ({
      code: "code-1",
      codeVerifier: "verifier-1",
    }));
    const app = {
      getAuthCodeUrl: vi.fn(),
      acquireTokenByCode: vi.fn(async () => ({
        accessToken: "token-1",
        expiresOn: new Date(now + 60 * 60 * 1000),
      })),
    };
    const provider = new AuthorizationCodeTokenProvider(CONFIG, { log: vi.fn() }, {
      app,
      authorize,
      now: () => now,
    });

    await expect(provider.getAccessToken()).resolves.toBe("token-1");
    await expect(provider.getAccessToken()).resolves.toBe("token-1");

    expect(authorize).toHaveBeenCalledTimes(1);
    expect(app.acquireTokenByCode).toHaveBeenCalledTimes(1);

    now += 56 * 60 * 1000;
    app.acquireTokenByCode.mockResolvedValueOnce({
      accessToken: "token-2",
      expiresOn: new Date(now + 60 * 60 * 1000),
    });

    await expect(provider.getAccessToken()).resolves.toBe("token-2");
    expect(authorize).toHaveBeenCalledTimes(2);
    expect(app.acquireTokenByCode).toHaveBeenCalledTimes(2);
  });

  it("shares an in-flight authorization request across concurrent token requests", async () => {
    const authorize = vi.fn(async () => ({
      code: "code-1",
      codeVerifier: "verifier-1",
    }));
    const app = {
      getAuthCodeUrl: vi.fn(),
      acquireTokenByCode: vi.fn(async () => ({
        accessToken: "token-1",
        expiresOn: new Date(Date.parse("2026-05-27T21:00:00.000Z")),
      })),
    };
    const provider = new AuthorizationCodeTokenProvider(CONFIG, { log: vi.fn() }, {
      app,
      authorize,
      now: () => Date.parse("2026-05-27T20:00:00.000Z"),
    });

    await expect(Promise.all([provider.getAccessToken(), provider.getAccessToken()])).resolves.toEqual([
      "token-1",
      "token-1",
    ]);

    expect(authorize).toHaveBeenCalledTimes(1);
    expect(app.acquireTokenByCode).toHaveBeenCalledTimes(1);
  });

  it("uses a persisted MSAL account silently before opening a browser", async () => {
    const authorize = vi.fn(async () => ({
      code: "code-1",
      codeVerifier: "verifier-1",
    }));
    const app = {
      getAuthCodeUrl: vi.fn(),
      acquireTokenByCode: vi.fn(),
      acquireTokenSilent: vi.fn(async () => ({
        accessToken: "silent-token",
        expiresOn: new Date(Date.parse("2026-05-27T21:00:00.000Z")),
      })),
      getTokenCache: vi.fn(() => ({
        getAllAccounts: vi.fn(async () => [
          {
            homeAccountId: "home-account-id",
            environment: "login.microsoftonline.com",
            tenantId: "organizations",
            username: "rex@example.com",
            localAccountId: "local-account-id",
          },
        ]),
      })),
    };
    const provider = new AuthorizationCodeTokenProvider(CONFIG, { log: vi.fn() }, {
      app,
      authorize,
      now: () => Date.parse("2026-05-27T20:00:00.000Z"),
    });

    await expect(provider.getAccessToken()).resolves.toBe("silent-token");

    expect(authorize).not.toHaveBeenCalled();
    expect(app.acquireTokenByCode).not.toHaveBeenCalled();
    expect(app.acquireTokenSilent).toHaveBeenCalledWith({
      account: expect.objectContaining({ homeAccountId: "home-account-id" }),
      scopes: CONFIG.scopes,
    });
  });
});
