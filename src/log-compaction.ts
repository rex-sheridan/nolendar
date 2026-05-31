export interface LogCompactionOptions {
  compactIds?: boolean;
}

const LONG_TOKEN_PATTERN = /[A-Za-z0-9_%=.~-]{32,}/g;
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DEFAULT_PREFIX_LENGTH = 10;
const DEFAULT_SUFFIX_LENGTH = 10;

export function compactLogIds(value: string, options: LogCompactionOptions = {}): string {
  if (!options.compactIds) {
    return value;
  }

  return value.replace(LONG_TOKEN_PATTERN, (token) => compactLogToken(token));
}

function compactLogToken(token: string): string {
  if (token.includes(",") || token.includes("%2C") || token.includes("%2c")) {
    return token;
  }

  const [keyPrefix, value] = splitKeyValueToken(token);
  const compactedValue = compactTokenValue(value);

  return compactedValue === value ? token : `${keyPrefix}${compactedValue}`;
}

function splitKeyValueToken(token: string): [string, string] {
  const equalsIndex = token.indexOf("=");

  if (equalsIndex <= 0 || equalsIndex === token.length - 1) {
    return ["", token];
  }

  return [token.slice(0, equalsIndex + 1), token.slice(equalsIndex + 1)];
}

function compactTokenValue(value: string): string {
  if (value.length < 32) {
    return value;
  }

  if (UUID_PATTERN.test(value)) {
    const [first, , , , last] = value.split("-");
    return `${first}...${last}`;
  }

  const suffixLength = Math.min(DEFAULT_SUFFIX_LENGTH, Math.floor((value.length - 3) / 2));
  const prefixLength = Math.min(DEFAULT_PREFIX_LENGTH, value.length - suffixLength - 3);

  if (prefixLength <= 0 || suffixLength <= 0 || prefixLength + suffixLength + 3 >= value.length) {
    return value;
  }

  return `${value.slice(0, prefixLength)}...${value.slice(-suffixLength)}`;
}
