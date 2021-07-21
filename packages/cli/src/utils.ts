// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import fs from "fs-extra";
import path from "path";
import { Options } from "yargs";
import chalk from "chalk";
import * as uuid from "uuid";
import * as dotenv from "dotenv";
import {
  OptionItem,
  Question,
  err,
  ok,
  Result,
  FxError,
  ConfigFolderName,
  getSingleOption,
  SingleSelectQuestion,
  MultiSelectQuestion,
  QTreeNode,
  Inputs,
  Platform,
  Colors,
} from "@microsoft/teamsfx-api";

import { ConfigNotFoundError, InvalidEnvFile, ReadFileError } from "./error";
import AzureAccountManager from "./commonlib/azureLogin";

type Json = { [_: string]: any };

export function getChoicesFromQTNodeQuestion(data: Question): string[] | undefined {
  const option = "staticOptions" in data ? data.staticOptions : undefined;
  if (option && option instanceof Array && option.length > 0) {
    if (typeof option[0] === "string") {
      return option as string[];
    } else {
      return (option as OptionItem[]).map((op) => (op.cliName ? op.cliName : op.id));
    }
  } else {
    return undefined;
  }
}

export function getSingleOptionString(
  q: SingleSelectQuestion | MultiSelectQuestion
): string | string[] {
  const singleOption = getSingleOption(q);
  if (q.returnObject) {
    if (q.type === "singleSelect") {
      return typeof singleOption === "string" ? singleOption : singleOption.id;
    } else {
      return [singleOption[0].id];
    }
  } else {
    return singleOption;
  }
}

export function toYargsOptions(data: Question): Options {
  const choices = getChoicesFromQTNodeQuestion(data);
  // if (choices && choices.length > 0 && data.default === undefined) {
  //   data.default = choices[0];
  // }

  const defaultValue = data.default;
  if (defaultValue && defaultValue instanceof Array && defaultValue.length > 0) {
    data.default = defaultValue.map((item) => item.toLocaleLowerCase());
  } else if (defaultValue && typeof defaultValue === "string") {
    data.default = defaultValue.toLocaleLowerCase();
  }
  if (data.default === undefined) {
    return {
      array: data.type === "multiSelect",
      description: data.title || "",
      choices: choices,
      hidden: !!(data as any).hide,
      global: false,
      type: "string",
    };
  }
  return {
    array: data.type === "multiSelect",
    description: data.title || "",
    default: data.default,
    choices: choices,
    hidden: !!(data as any).hide,
    global: false,
    type: "string",
  };
}

export function flattenNodes(node: QTreeNode): QTreeNode[] {
  const nodeCopy = Object.assign({}, node);
  const children = (nodeCopy.children || []).concat([]);
  nodeCopy.children = undefined;
  return [nodeCopy].concat(...children.map((nd) => flattenNodes(nd)));
}

export async function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function getActiveEnv(): string {
  return "default";
}

export function getConfigPath(projectFolder: string, fileName: string): string {
  return path.resolve(projectFolder, `.${ConfigFolderName}`, fileName);
}

export function getEnvFilePath(projectFolder: string) {
  return getConfigPath(projectFolder, `env.${getActiveEnv()}.json`);
}

export async function readEnvJsonFile(projectFolder: string): Promise<Result<Json, FxError>> {
  const filePath = getEnvFilePath(projectFolder);
  if (!fs.existsSync(filePath)) {
    return err(ConfigNotFoundError(filePath));
  }
  try {
    const config = await fs.readJson(filePath);
    return ok(config);
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export function readEnvJsonFileSync(projectFolder: string): Result<Json, FxError> {
  const filePath = getEnvFilePath(projectFolder);
  if (!fs.existsSync(filePath)) {
    return err(ConfigNotFoundError(filePath));
  }
  try {
    const config = fs.readJsonSync(filePath);
    return ok(config);
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export async function readProjectSecrets(
  projectFolder: string
): Promise<Result<dotenv.DotenvParseOutput, FxError>> {
  const secretFile = getConfigPath(projectFolder, `${getActiveEnv()}.userdata`);
  if (!fs.existsSync(secretFile)) {
    return err(ConfigNotFoundError(secretFile));
  }
  try {
    const secretData = await fs.readFile(secretFile);
    return ok(dotenv.parse(secretData));
  } catch (e) {
    return err(ReadFileError(e));
  }
}

export function writeSecretToFile(secrets: dotenv.DotenvParseOutput, rootFolder: string): void {
  const secretFile = `${rootFolder}/.${ConfigFolderName}/${getActiveEnv()}.userdata`;
  const array: string[] = [];
  for (const secretKey of Object.keys(secrets)) {
    const secretValue = secrets[secretKey];
    array.push(`${secretKey}=${secretValue}`);
  }
  fs.writeFileSync(secretFile, array.join("\n"));
}

export async function getSolutionPropertyFromEnvFile(
  projectFolder: string,
  propertyName: string
): Promise<Result<any, FxError>> {
  const result = await readEnvJsonFile(projectFolder);
  if (result.isErr()) {
    return err(result.error);
  }
  const env = result.value;
  if ("solution" in env) {
    return ok(env.solution[propertyName]);
  } else {
    return err(
      InvalidEnvFile(
        `The property \`solution\` does not exist in the project's env file.`,
        getEnvFilePath(projectFolder)
      )
    );
  }
}

export async function getSubscriptionIdFromEnvFile(
  rootFolder: string
): Promise<string | undefined> {
  const result = await getSolutionPropertyFromEnvFile(rootFolder, "subscriptionId");
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
}

export async function setSubscriptionId(
  subscriptionId?: string,
  rootFolder = "./"
): Promise<Result<null, FxError>> {
  if (subscriptionId) {
    const result = await readEnvJsonFile(rootFolder);
    if (result.isErr()) {
      return err(result.error);
    }

    AzureAccountManager.setRootPath(rootFolder);
    await AzureAccountManager.setSubscription(subscriptionId);
    const subs = await AzureAccountManager.listSubscriptions();
    const sub = subs.find((sub) => sub.subscriptionId === subscriptionId);

    const configJson = result.value;
    configJson["solution"].subscriptionId = sub?.subscriptionId;
    configJson["solution"].tenantId = sub?.tenantId;
    await fs.writeFile(getEnvFilePath(rootFolder), JSON.stringify(configJson, null, 4));
  }
  return ok(null);
}

export function isWorkspaceSupported(workspace: string): boolean {
  const p = workspace;

  const checklist: string[] = [
    p,
    `${p}/package.json`,
    `${p}/.${ConfigFolderName}`,
    `${p}/.${ConfigFolderName}/settings.json`,
    `${getEnvFilePath(p)}`,
  ];

  for (const fp of checklist) {
    if (!fs.existsSync(path.resolve(fp))) {
      return false;
    }
  }
  return true;
}

export function getTeamsAppId(rootfolder: string | undefined): any {
  if (!rootfolder) {
    return undefined;
  }

  if (isWorkspaceSupported(rootfolder)) {
    const result = readEnvJsonFileSync(rootfolder);
    if (result.isErr()) {
      throw result.error;
    }
    return result.value.solution.remoteTeamsAppId;
  }

  return undefined;
}

export function getLocalTeamsAppId(rootfolder: string | undefined): any {
  if (!rootfolder) {
    return undefined;
  }

  if (isWorkspaceSupported(rootfolder)) {
    const result = readEnvJsonFileSync(rootfolder);
    if (result.isErr()) {
      throw result.error;
    }
    return result.value.solution.localDebugTeamsAppId;
  }

  return undefined;
}

export function getSystemInputs(projectPath?: string): Inputs {
  const systemInputs: Inputs = {
    platform: Platform.CLI,
    projectPath: projectPath,
    correlationId: uuid.v4(),
  };
  return systemInputs;
}

export function argsToInputs(
  params: { [_: string]: Options },
  args: { [argName: string]: string | string[] }
): Inputs {
  const inputs = getSystemInputs();
  for (const name in params) {
    if (name.endsWith("folder") && args[name]) {
      inputs[name] = path.resolve(args[name] as string);
    } else {
      inputs[name] = args[name];
    }
  }
  const rootFolder = path.resolve((inputs["folder"] as string) || "./");
  delete inputs["folder"];
  inputs.projectPath = rootFolder;
  return inputs;
}

export function getColorizedString(message: Array<{ content: string; color: Colors }>): string {
  // Color support is automatically detected by chalk
  const colorizedMessage = message
    .map((item) => {
      switch (item.color) {
        case Colors.BRIGHT_WHITE:
          return chalk.whiteBright(item.content);
        case Colors.WHITE:
          return chalk.white(item.content);
        case Colors.BRIGHT_MAGENTA:
          return chalk.magentaBright(item.content);
        case Colors.BRIGHT_GREEN:
          return chalk.greenBright(item.content);
        case Colors.BRIGHT_RED:
          return chalk.redBright(item.content);
        case Colors.BRIGHT_YELLOW:
          return chalk.yellowBright(item.content);
        case Colors.BRIGHT_CYAN:
          return chalk.cyanBright.underline(item.content);
        default:
          return item.content;
      }
    })
    .join("");
  return colorizedMessage + (process.stdout.isTTY ? "\u00A0\u001B[K" : "");
}

/**
 * Shows in `teamsfx -v`.
 * @returns the version of teamsfx-cli.
 */
export function getVersion(): string {
  const pkgPath = path.resolve(__dirname, "..", "package.json");
  const pkgContent = fs.readJsonSync(pkgPath);
  return pkgContent.version;
}
