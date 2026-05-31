import { readFile } from "node:fs/promises";

export class EnvFileError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "EnvFileError";
  }
}

export async function loadLocalEnvFile(envPath = ".env", env: NodeJS.ProcessEnv = process.env): Promise<void> {
  let raw: string;

  try {
    raw = await readFile(envPath, "utf8");
  } catch (error) {
    if (isMissingFileError(error)) {
      return;
    }

    throw error;
  }

  for (const [key, value] of parseEnvFile(raw, envPath)) {
    if (env[key] === undefined) {
      env[key] = value;
    }
  }
}

export function parseEnvFile(raw: string, envPath = ".env"): Map<string, string> {
  const values = new Map<string, string>();
  const lines = raw.replace(/^\uFEFF/, "").split(/\r?\n/);

  lines.forEach((line, index) => {
    const trimmed = line.trim();

    if (trimmed === "" || trimmed.startsWith("#")) {
      return;
    }

    const assignment = trimmed.startsWith("export ") ? trimmed.slice("export ".length).trimStart() : trimmed;
    const separatorIndex = assignment.indexOf("=");

    if (separatorIndex <= 0) {
      throw new EnvFileError(`${envPath}:${index + 1} must be a KEY=value assignment.`);
    }

    const key = assignment.slice(0, separatorIndex).trim();
    const value = assignment.slice(separatorIndex + 1).trim();

    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) {
      throw new EnvFileError(`${envPath}:${index + 1} has an invalid environment variable name.`);
    }

    values.set(key, parseEnvValue(value));
  });

  return values;
}

function parseEnvValue(value: string): string {
  if (value.startsWith('"') && value.endsWith('"')) {
    return value.slice(1, -1).replace(/\\n/g, "\n").replace(/\\"/g, '"').replace(/\\\\/g, "\\");
  }

  if (value.startsWith("'") && value.endsWith("'")) {
    return value.slice(1, -1);
  }

  const commentIndex = value.search(/\s#/);
  return (commentIndex === -1 ? value : value.slice(0, commentIndex)).trimEnd();
}

function isMissingFileError(error: unknown): boolean {
  return typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT";
}
