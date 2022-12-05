// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  Result,
  err,
  ok,
  TeamsAppManifest,
  UserCancelError,
  Platform,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import AdmZip from "adm-zip";
import { hooks } from "@feathersjs/hooks/lib";
import { StepDriver, ExecutionResult } from "../interface/stepDriver";
import { DriverContext } from "../interface/commonArgs";
import { WrapDriverContext } from "../util/wrapUtil";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { PublishAppPackageArgs } from "./interfaces/PublishAppPackageArgs";
import { AppStudioClient } from "../../resource/appManifest/appStudioClient";
import { Constants } from "../../resource/appManifest/constants";
import { AppStudioResultFactory } from "../../resource/appManifest/results";
import { TelemetryUtils } from "../../resource/appManifest/utils/telemetry";
import { AppStudioError } from "../../resource/appManifest/errors";
import { TelemetryPropertyKey } from "../../resource/appManifest/utils/telemetry";
import { AppStudioScopes } from "../../../common/tools";
import { getLocalizedString } from "../../../common/localizeUtils";
import { Service } from "typedi";
import { getAbsolutePath } from "../../utils/common";

const actionName = "teamsApp/publishAppPackage";

const outputKeys = {
  publishedAppId: "TEAMS_APP_PUBLISHED_APP_ID",
};

@Service(actionName)
export class PublishAppPackageDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.publishDriver");
  public async run(
    args: PublishAppPackageArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.publish(args, wrapContext);

    console.log("Summaries");
    wrapContext.summaries.forEach((value) => console.log(value));

    return res;
  }

  public async execute(
    args: PublishAppPackageArgs,
    context: DriverContext
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.publish(args, wrapContext);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async publish(
    args: PublishAppPackageArgs,
    context: WrapDriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    TelemetryUtils.init(context);
    const progressHandler = context.ui?.createProgressBar(
      getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppTitle"),
      2
    );
    progressHandler?.start();

    const appPackagePath = getAbsolutePath(args.appPackagePath, context.projectPath);
    if (!(await fs.pathExists(appPackagePath))) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.FileNotFoundError.name,
          AppStudioError.FileNotFoundError.message(args.appPackagePath),
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    }
    const archivedFile = await fs.readFile(appPackagePath);

    const zipEntries = new AdmZip(archivedFile).getEntries();

    const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
    if (!manifestFile) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.FileNotFoundError.name,
          AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE),
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    }
    const manifestString = manifestFile.getData().toString();
    const manifest = JSON.parse(manifestString) as TeamsAppManifest;

    // manifest.id === externalID
    const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes,
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }

    let result;
    const telemetryProps: { [key: string]: string } = {};

    const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep1");
    progressHandler?.next(message);
    context.addSummary(message);

    const existApp = await AppStudioClient.getAppByTeamsAppId(manifest.id, appStudioTokenRes.value);
    if (existApp) {
      context.addSummary(
        getLocalizedString("driver.teamsApp.summary.publishTeamsAppExists", manifest.id)
      );
      let executePublishUpdate = false;
      let description = getLocalizedString(
        "plugins.appstudio.pubWarn",
        existApp.displayName,
        existApp.publishingState
      );
      if (existApp.lastModifiedDateTime) {
        description =
          description +
          getLocalizedString(
            "plugins.appstudio.lastModified",
            existApp.lastModifiedDateTime?.toLocaleString()
          );
      }
      description = description + getLocalizedString("plugins.appstudio.updatePublihsedAppConfirm");
      const confirm = getLocalizedString("core.option.confirm");
      const res = await context.ui?.showMessage("warn", description, true, confirm);
      if (res?.isOk() && res.value === confirm) executePublishUpdate = true;

      if (executePublishUpdate) {
        const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.1");
        progressHandler?.next(message);
        context.addSummary(message);
        const appId = await AppStudioClient.publishTeamsAppUpdate(
          manifest.id,
          archivedFile,
          appStudioTokenRes.value
        );
        result = new Map([[outputKeys.publishedAppId, appId]]);
        // TODO: how to send telemetry with own properties
        telemetryProps[TelemetryPropertyKey.updateExistingApp] = "true";
      } else {
        progressHandler?.end(true);
        return err(UserCancelError);
      }
    } else {
      context.addSummary(
        getLocalizedString("driver.teamsApp.summary.publishTeamsAppNotExists", manifest.id)
      );
      const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.2");
      progressHandler?.next(message);
      context.addSummary(message);
      const appId = await AppStudioClient.publishTeamsApp(
        manifest.id,
        archivedFile,
        appStudioTokenRes.value
      );
      result = new Map([[outputKeys.publishedAppId, appId]]);
      telemetryProps[TelemetryPropertyKey.updateExistingApp] = "false";
    }

    progressHandler?.end(true);

    context.logProvider.info(`Publish success!`);
    context.addSummary(
      getLocalizedString("driver.teamsApp.summary.publishTeamsAppSuccess", manifest.id)
    );
    if (context.platform === Platform.CLI) {
      const msg = getLocalizedString(
        "plugins.appstudio.publishSucceedNotice.cli",
        manifest.name.short,
        Constants.TEAMS_ADMIN_PORTAL,
        Constants.TEAMS_MANAGE_APP_DOC
      );
      context.ui?.showMessage("info", msg, false);
    } else {
      const msg = getLocalizedString(
        "plugins.appstudio.publishSucceedNotice",
        manifest.name.short,
        Constants.TEAMS_MANAGE_APP_DOC
      );
      const adminPortal = getLocalizedString("plugins.appstudio.adminPortal");
      context.ui?.showMessage("info", msg, false, adminPortal).then((value) => {
        if (value.isOk() && value.value === adminPortal) {
          context.ui?.openUrl(Constants.TEAMS_ADMIN_PORTAL);
        }
      });
    }
    return ok(result);
  }
}
