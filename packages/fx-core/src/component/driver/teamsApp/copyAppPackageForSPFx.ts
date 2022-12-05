// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { Result, FxError, TeamsAppManifest } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import { Service } from "typedi";
import { getLocalizedString } from "../../../common/localizeUtils";
import { Constants } from "../../resource/appManifest/constants";
import { AppStudioError } from "../../resource/appManifest/errors";
import { AppStudioResultFactory } from "../../resource/appManifest/results";
import { asFactory, asString, wrapRun } from "../../utils/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { CopyAppPackageForSPFxArgs } from "./interfaces/CopyAppPackageForSPFxArgs";

const actionName = "teamsApp/copyAppPackageForSPFx";

@Service(actionName)
export class CopyAppPackageForSPFxDriver implements StepDriver {
  public readonly description = getLocalizedString(
    "driver.teamsApp.description.copyAppPackageForSPFxDriver"
  );

  private readonly EmptyMap = new Map<string, string>();

  private asCopyAppPackageArgs = asFactory<CopyAppPackageForSPFxArgs>({
    appPackagePath: asString,
    spfxFolder: asString,
  });

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async run(
    args: CopyAppPackageForSPFxArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    return wrapRun(() => this.copy(args, wrapContext));
  }

  public async execute(
    args: CopyAppPackageForSPFxArgs,
    ctx: DriverContext
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(ctx, actionName, actionName);
    const result = await this.run(args, wrapContext);
    return {
      result,
      summaries: wrapContext.summaries,
    };
  }

  public async copy(
    args: CopyAppPackageForSPFxArgs,
    context: WrapDriverContext
  ): Promise<Map<string, string>> {
    const copyAppPackageArgs = this.asCopyAppPackageArgs(args);
    const appPackagePath = path.isAbsolute(copyAppPackageArgs.appPackagePath)
      ? copyAppPackageArgs.appPackagePath
      : path.join(context.projectPath, copyAppPackageArgs.appPackagePath);
    if (!(await fs.pathExists(appPackagePath))) {
      throw AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(appPackagePath)
      );
    }
    const pictures = await this.getIcons(appPackagePath);
    const spfxFolder = path.isAbsolute(copyAppPackageArgs.spfxFolder)
      ? copyAppPackageArgs.spfxFolder
      : path.join(context.projectPath, copyAppPackageArgs.spfxFolder);
    const spfxTeamsPath = `${spfxFolder}/teams`;
    await fs.copyFile(appPackagePath, path.join(spfxTeamsPath, "TeamsSPFxApp.zip"));
    context.addSummary(
      getLocalizedString(
        "driver.teamsApp.summary.copyAppPackageSuccess",
        appPackagePath,
        path.join(spfxTeamsPath, "TeamsSPFxApp.zip")
      )
    );

    let replacedIcons = 0;
    for (const file of await fs.readdir(spfxTeamsPath)) {
      if (file.endsWith("color.png") && pictures.color) {
        await fs.writeFile(path.join(spfxTeamsPath, file), pictures.color);
        replacedIcons++;
      } else if (file.endsWith("outline.png") && pictures.outline) {
        await fs.writeFile(path.join(spfxTeamsPath, file), pictures.outline);
        replacedIcons++;
      }
    }
    if (replacedIcons > 0) {
      context.addSummary(
        getLocalizedString("driver.teamsApp.summary.copyIconSuccess", replacedIcons, spfxTeamsPath)
      );
    }
    return this.EmptyMap;
  }

  public async getIcons(appPackagePath: string): Promise<IIcons> {
    const archivedFile = await fs.readFile(appPackagePath);
    const zipEntries = new AdmZip(archivedFile).getEntries();
    const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
    if (!manifestFile) {
      throw AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE)
      );
    }
    const manifestString = manifestFile.getData().toString();
    const manifest = JSON.parse(manifestString) as TeamsAppManifest;

    const colorFile =
      manifest.icons.color && !manifest.icons.color.startsWith("https://")
        ? zipEntries.find((x) => x.entryName.includes("color.png"))
        : undefined;
    const outlineFile =
      manifest.icons.outline && !manifest.icons.outline.startsWith("https://")
        ? zipEntries.find((x) => x.entryName.includes("outline.png"))
        : undefined;
    return { color: colorFile?.getData(), outline: outlineFile?.getData() };
  }
}

interface IIcons {
  color?: Buffer;
  outline?: Buffer;
}
