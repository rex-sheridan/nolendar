import { mkdir, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";

export const DEFAULT_MICROSOFT_TOKEN_CACHE_PATH = path.join(os.homedir(), ".nolendar", "msal-cache.json");

export class FileTokenCachePlugin implements ICachePlugin {
  constructor(private readonly cachePath = DEFAULT_MICROSOFT_TOKEN_CACHE_PATH) {}

  async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    try {
      cacheContext.tokenCache.deserialize(await readFile(this.cachePath, "utf8"));
    } catch (error) {
      if (isMissingFileError(error)) {
        return;
      }

      throw error;
    }
  }

  async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    if (!cacheContext.cacheHasChanged) {
      return;
    }

    await mkdir(path.dirname(this.cachePath), { recursive: true, mode: 0o700 });
    await writeFile(this.cachePath, cacheContext.tokenCache.serialize(), { encoding: "utf8", mode: 0o600 });
  }
}

function isMissingFileError(error: unknown): boolean {
  return typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT";
}
