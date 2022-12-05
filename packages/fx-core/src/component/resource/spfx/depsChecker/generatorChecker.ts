// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import * as path from "path";
import * as os from "os";
import {
  ConfigFolderName,
  ContextV3,
  err,
  FxError,
  LogProvider,
  ok,
  PluginContext,
  Result,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import { DependencyChecker, DependencyInfo } from "./dependencyChecker";
import { telemetryHelper } from "../utils/telemetry-helper";
import { TelemetryEvents, TelemetryProperty } from "../utils/telemetryEvents";
import { DependencyValidateError, NpmInstallError } from "../error";
import { cpUtils } from "../../../../common/deps-checker/util/cpUtils";
import { Constants } from "../utils/constants";

const name = "@microsoft/generator-sharepoint";
const supportedVersion = Constants.SPFX_VERSION;
const displayName = `${name}@${supportedVersion}`;
const timeout = 6 * 60 * 1000;

export class GeneratorChecker implements DependencyChecker {
  private readonly _logger: LogProvider;

  constructor(logger: LogProvider) {
    this._logger = logger;
  }

  public static getDependencyInfo(): DependencyInfo {
    return {
      supportedVersion: supportedVersion,
      displayName: displayName,
    };
  }

  public async ensureDependency(ctx: PluginContext | ContextV3): Promise<Result<boolean, FxError>> {
    telemetryHelper.sendSuccessEvent(ctx, TelemetryEvents.EnsureSharepointGeneratorStart);
    try {
      if (!(await this.isInstalled())) {
        this._logger.info(`${displayName} not found, installing ...`);
        await this.install();
        this._logger.info(`Successfully installed ${displayName}`);
      }
      telemetryHelper.sendSuccessEvent(ctx, TelemetryEvents.EnsureSharepointGenerator);
    } catch (error) {
      telemetryHelper.sendErrorEvent(
        ctx,
        TelemetryEvents.EnsureSharepointGenerator,
        error as UserError | SystemError,
        {
          [TelemetryProperty.EnsureSharepointGeneratorReason]: (error as UserError | SystemError)
            .name,
        }
      );
      await this._logger.error(`Failed to install ${name}, error = '${error}'`);
      return err(error as UserError | SystemError);
    }

    return ok(true);
  }

  public async isInstalled(): Promise<boolean> {
    let isVersionSupported = false,
      hasSentinel = false;
    try {
      const generatorVersion = await this.queryVersion();
      isVersionSupported = generatorVersion !== undefined && supportedVersion === generatorVersion;
      hasSentinel = await fs.pathExists(this.getSentinelPath());
    } catch (error) {
      return false;
    }
    return isVersionSupported && hasSentinel;
  }

  public async install(): Promise<void> {
    this._logger.info("Start installing...");
    await this.cleanup();
    await this.installGenerator();

    this._logger.info("Validating package...");
    if (!(await this.validate())) {
      this._logger.debug(`Failed to validate ${name}, cleaning up...`);
      await this.cleanup();
      throw DependencyValidateError(name);
    }
  }

  public getSpGeneratorPath(): string {
    return `"${path.join(
      this.getDefaultInstallPath(),
      "node_modules",
      "@microsoft",
      "generator-sharepoint",
      "lib",
      "generators",
      "app",
      "index.js"
    )}"`;
  }

  private async validate(): Promise<boolean> {
    return await this.isInstalled();
  }

  private getDefaultInstallPath(): string {
    return path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "spGenerator");
  }

  private getSentinelPath(): string {
    return path.join(os.homedir(), `.${ConfigFolderName}`, "spGenerator-sentinel");
  }

  private async queryVersion(): Promise<string | undefined> {
    const packagePath = path.join(
      this.getDefaultInstallPath(),
      "node_modules",
      "@microsoft",
      "generator-sharepoint",
      "package.json"
    );
    if (await fs.pathExists(packagePath)) {
      const packageJson = await fs.readJson(packagePath);
      return packageJson.version ?? undefined;
    }
    return undefined;
  }

  private async cleanup(): Promise<void> {
    try {
      const legacyDirectory = path.join(os.homedir(), `.${ConfigFolderName}`, "bin", "spfx");
      if (fs.existsSync(legacyDirectory)) {
        await fs.emptyDir(legacyDirectory);
        await fs.rmdir(legacyDirectory);
      }

      await fs.emptyDir(this.getDefaultInstallPath());
      await fs.remove(this.getSentinelPath());
    } catch (err) {
      await this._logger.error(
        `Failed to clean up path: ${this.getDefaultInstallPath()}, error: ${err}`
      );
    }
  }

  private async installGenerator(): Promise<void> {
    try {
      await fs.ensureDir(path.join(this.getDefaultInstallPath(), "node_modules"));
      await cpUtils.executeCommand(
        undefined,
        this._logger,
        { timeout: timeout, shell: false },
        this.getExecCommand("npm"),
        "install",
        `${name}@${supportedVersion}`,
        "--prefix",
        `${this.getDefaultInstallPath()}`,
        "--no-audit",
        "--global-style"
      );

      await fs.ensureFile(this.getSentinelPath());
    } catch (error) {
      this._logger.error(`Failed to execute npm install ${displayName}`);
      throw NpmInstallError(error as Error);
    }
  }

  private getExecCommand(command: string): string {
    return this.isWindows() ? `${command}.cmd` : command;
  }

  private isWindows(): boolean {
    return os.type() === "Windows_NT";
  }
}
