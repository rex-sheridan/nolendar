import path from "node:path";
import { describe, expect, it } from "vitest";

import {
  defaultConfigFilePath,
  defaultDataDirectory,
  defaultEnvFilePath,
  defaultStateFilePath,
} from "../src/xdg.js";

describe("XDG paths", () => {
  it("uses the XDG base directories when they are configured", () => {
    const env = {
      XDG_CONFIG_HOME: "/xdg/config",
      XDG_DATA_HOME: "/xdg/data",
      XDG_STATE_HOME: "/xdg/state",
    };

    expect(defaultConfigFilePath(env, "/home/rex")).toBe("/xdg/config/nolendar/config.yml");
    expect(defaultEnvFilePath(env, "/home/rex")).toBe("/xdg/config/nolendar/env");
    expect(defaultDataDirectory(env, "/home/rex")).toBe("/xdg/data/nolendar");
    expect(defaultStateFilePath(env, "/home/rex")).toBe("/xdg/state/nolendar/state.json");
  });

  it("falls back to the specification defaults under the user's home directory", () => {
    expect(defaultConfigFilePath({}, "/home/rex")).toBe("/home/rex/.config/nolendar/config.yml");
    expect(defaultDataDirectory({}, "/home/rex")).toBe("/home/rex/.local/share/nolendar");
    expect(defaultStateFilePath({}, "/home/rex")).toBe("/home/rex/.local/state/nolendar/state.json");
  });

  it("ignores relative XDG base directory values as required by the specification", () => {
    const env = {
      XDG_CONFIG_HOME: "relative/config",
      XDG_DATA_HOME: "relative/data",
      XDG_STATE_HOME: "relative/state",
    };

    expect(defaultConfigFilePath(env, "/home/rex")).toBe(path.join("/home/rex", ".config/nolendar/config.yml"));
    expect(defaultDataDirectory(env, "/home/rex")).toBe(path.join("/home/rex", ".local/share/nolendar"));
    expect(defaultStateFilePath(env, "/home/rex")).toBe(path.join("/home/rex", ".local/state/nolendar/state.json"));
  });
});
