import os from "node:os";
import path from "node:path";

const APPLICATION_DIRECTORY = "nolendar";

export function defaultConfigDirectory(env = process.env, homeDir = os.homedir()): string {
  return path.join(resolveBaseDirectory(env.XDG_CONFIG_HOME, path.join(homeDir, ".config")), APPLICATION_DIRECTORY);
}

export function defaultDataDirectory(env = process.env, homeDir = os.homedir()): string {
  return path.join(resolveBaseDirectory(env.XDG_DATA_HOME, path.join(homeDir, ".local", "share")), APPLICATION_DIRECTORY);
}

export function defaultStateDirectory(env = process.env, homeDir = os.homedir()): string {
  return path.join(resolveBaseDirectory(env.XDG_STATE_HOME, path.join(homeDir, ".local", "state")), APPLICATION_DIRECTORY);
}

export function defaultConfigFilePath(env = process.env, homeDir = os.homedir()): string {
  return path.join(defaultConfigDirectory(env, homeDir), "config.yml");
}

export function defaultEnvFilePath(env = process.env, homeDir = os.homedir()): string {
  return path.join(defaultConfigDirectory(env, homeDir), "env");
}

export function defaultMicrosoftTokenCachePath(env = process.env, homeDir = os.homedir()): string {
  return path.join(defaultDataDirectory(env, homeDir), "msal-cache.json");
}

export function defaultStateFilePath(env = process.env, homeDir = os.homedir()): string {
  return path.join(defaultStateDirectory(env, homeDir), "state.json");
}

function resolveBaseDirectory(candidate: string | undefined, fallback: string): string {
  const value = candidate?.trim();
  return value && path.isAbsolute(value) ? value : fallback;
}
