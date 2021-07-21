// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { OptionItem } from "./qm";
import { Platform, Stage, VsCodeEnv } from "./constants";

export type Json = Record<string, unknown>;

export type ConfigValue = any;

export type PluginIdentity = string;

export type PluginConfig = ConfigMap;
export type ReadonlyPluginConfig = ReadonlyMap<string, ConfigValue>;

export type SolutionConfig = Map<PluginIdentity, PluginConfig>;
export type ReadonlySolutionConfig = ReadonlyMap<PluginIdentity, ReadonlyPluginConfig>;

export class ConfigMap extends Map<string, ConfigValue> {
  getString(k: string, defaultValue?: string): string | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return String(v);
  }
  getBoolean(k: string, defaultValue?: boolean): boolean | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return Boolean(v);
  }
  getNumber(k: string, defaultValue?: number): number | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return Number(v);
  }
  getStringArray(k: string, defaultValue?: string[]): string[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as string[];
  }
  getNumberArray(k: string, defaultValue?: number[]): number[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as number[];
  }
  getBooleanArray(k: string, defaultValue?: boolean[]): boolean[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as boolean[];
  }
  getOptionItem(k: string, defaultValue?: OptionItem): OptionItem | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as OptionItem;
  }
  getOptionItemArray(k: string, defaultValue?: OptionItem[]): OptionItem[] | undefined {
    const v = super.get(k);
    if (!v) return defaultValue;
    return v as OptionItem[];
  }

  toJSON(): Dict<unknown> {
    const out: Dict<unknown> = {};
    for (const entry of super.entries()) {
      out[entry[0]] = entry[1];
    }
    return out;
  }

  public static fromJSON(obj?: Dict<unknown>): ConfigMap | undefined {
    if (!obj) return undefined;
    const map = new ConfigMap();
    for (const entry of Object.entries(obj)) {
      map.set(entry[0], entry[1]);
    }
    return map;
  }
  // constructor(){
  //     super();
  //     Object.setPrototypeOf(this, ConfigMap.prototype);
  // }
  constructor(entries?: readonly (readonly [string, ConfigValue])[] | null) {
    super(entries);
    Object.setPrototypeOf(this, ConfigMap.prototype);
  }
}

// eslint-disable-next-line @typescript-eslint/ban-types
export type Void = {};
export const Void = {};

export interface Dict<T> {
  [key: string]: T | undefined;
}

export type ResourceTemplate = Dict<ConfigValue>;

export type ResourceTemplates = {
  [k: string]: ResourceTemplate | undefined;
};

export type ResourceConfig = ResourceTemplate;

export type ResourceConfigs = ResourceTemplates;

export type ReadonlyResourceConfig = Readonly<ResourceConfig>;

export type ReadonlyResourceConfigs = Readonly<{
  [k: string]: ReadonlyResourceConfig | undefined;
}>;

/**
 * environment meta data
 */
export interface EnvMeta {
  name: string;
  local: boolean;
  sideloading: boolean;
}

export type EnvConfig = Dict<string>;

/**
 * project static settings
 */
export interface ProjectSettings {
  appName: string;
  projectId: string;
  currentEnv?: string;
  solutionSettings?: SolutionSettings;
}

/**
 * solution settings
 */
export interface SolutionSettings extends Dict<ConfigValue> {
  name: string;
  version: string;
}

export interface AzureSolutionSettings extends SolutionSettings {
  hostType: string;
  capabilities: string[];
  azureResources: string[];
  activeResourcePlugins: string[];
}

/**
 * project dynamic states
 */
export interface ProjectStates {
  solution: Dict<ConfigValue>;
  resources: {
    [k: string]: Dict<ConfigValue>;
  };
}

export interface Inputs extends Json {
  projectPath?: string;
  platform: Platform;
  stage?: Stage;
  vscodeEnv?: VsCodeEnv;
  ignoreLock?: boolean;
  ignoreTypeCheck?: boolean;
  ignoreConfigPersist?: boolean;
  correlationId?: string;
}

export interface ProjectConfig {
  settings?: ProjectSettings;
  config?: SolutionConfig;
}
