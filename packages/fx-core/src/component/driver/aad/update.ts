// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { DriverContext } from "../interface/commonArgs";
import { UpdateAadAppArgs } from "./interface/updateAadAppArgs";
import { Service } from "typedi";
import { InvalidParameterUserError } from "./error/invalidParameterUserError";
import { UpdateAadAppOutput } from "./interface/updateAadAppOutput";
import { ProgressBarSetting } from "./interface/progressBarSetting";
import * as fs from "fs-extra";
import * as path from "path";
import { AadAppClient } from "./utility/aadAppClient";
import axios from "axios";
import { SystemError, UserError, ok, err, FxError, Result } from "@microsoft/teamsfx-api";
import { UnhandledSystemError, UnhandledUserError } from "./error/unhandledError";
import { getUuid } from "../../../common/tools";
import { expandEnvironmentVariable, getEnvironmentVariables } from "../../utils/common";
import { AadManifestHelper } from "../../resource/aadApp/utils/aadManifestHelper";
import { AADManifest } from "../../resource/aadApp/interfaces/AADManifest";
import { MissingFieldInManifestUserError } from "./error/invalidFieldInManifestError";
import isUUID from "validator/lib/isUUID";
import { hooks } from "@feathersjs/hooks/lib";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { getLocalizedString } from "../../../common/localizeUtils";
import { logMessageKeys, descriptionMessageKeys } from "./utility/constants";
import { MissingEnvUserError } from "./error/missingEnvError";

const actionName = "aadApp/update"; // DO NOT MODIFY the name
const helpLink = "https://aka.ms/teamsfx-actions/aadapp-update";
const driverConstants = {
  generateManifestFailedMessageKey: "driver.aadApp.error.generateManifestFailed",
};

// logic from src\component\resource\aadApp\aadAppManifestManager.ts
@Service(actionName) // DO NOT MODIFY the service name
export class UpdateAadAppDriver implements StepDriver {
  description = getLocalizedString(descriptionMessageKeys.update);

  public async run(
    args: UpdateAadAppArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const result = await this.execute(args, context);
    return result.result;
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(args: UpdateAadAppArgs, context: DriverContext): Promise<ExecutionResult> {
    const progressBarSettings = this.getProgressBarSetting();
    const progressHandler = context.ui?.createProgressBar(
      progressBarSettings.title,
      progressBarSettings.stepMessages.length
    );
    const summaries: string[] = [];

    try {
      await progressHandler?.start();
      await progressHandler?.next(progressBarSettings.stepMessages.shift());
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));

      this.validateArgs(args);
      const aadAppClient = new AadAppClient(context.m365TokenProvider);
      const state = this.loadCurrentState();

      const manifestAbsolutePath = this.getAbsolutePath(
        args.manifestTemplatePath,
        context.projectPath
      );
      const manifest = await this.loadManifest(manifestAbsolutePath, state);
      const warningMessage = AadManifestHelper.validateManifest(manifest);
      if (warningMessage) {
        warningMessage.split("\n").forEach((warning) => {
          context.logProvider?.warning(warning);
        });
      }

      if (!manifest.id || !isUUID(manifest.id)) {
        throw new MissingFieldInManifestUserError(actionName, "id", helpLink);
      }

      // Output actual manifest to project folder first for better troubleshooting experience
      const outputFileAbsolutePath = this.getAbsolutePath(args.outputFilePath, context.projectPath);
      await fs.ensureDir(path.dirname(outputFileAbsolutePath));
      await fs.writeFile(outputFileAbsolutePath, JSON.stringify(manifest, null, 4), "utf8");
      context.logProvider?.info(
        getLocalizedString(logMessageKeys.outputAadAppManifest, outputFileAbsolutePath)
      );

      // MS Graph API does not allow adding new OAuth permissions and pre authorize it within one request
      // So split update AAD app to two requests:
      // 1. If there's preAuthorizedApplications, remove it temporary and update AAD app to create possible new permission
      if (manifest.preAuthorizedApplications && manifest.preAuthorizedApplications.length > 0) {
        const preAuthorizedApplications = manifest.preAuthorizedApplications;
        manifest.preAuthorizedApplications = [];
        await aadAppClient.updateAadApp(manifest);
        manifest.preAuthorizedApplications = preAuthorizedApplications;
      }
      // 2. Update AAD app again with full manifest to set preAuthorizedApplications
      await aadAppClient.updateAadApp(manifest);
      const summary = getLocalizedString(
        logMessageKeys.successUpdateAadAppManifest,
        args.manifestTemplatePath,
        manifest.id
      );
      context.logProvider?.info(summary);
      summaries.push(summary);

      context.logProvider?.info(
        getLocalizedString(logMessageKeys.successExecuteDriver, actionName)
      );
      await progressHandler?.end(true);

      return {
        result: ok(
          new Map(
            Object.entries(state) // convert each property to Map item
              .filter((item) => item[1] && item[1] !== "") // do not return Map item that is empty
          )
        ),
        summaries: summaries,
      };
    } catch (error) {
      await progressHandler?.end(false);
      if (error instanceof UserError || error instanceof SystemError) {
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, error.displayMessage)
        );
        return {
          result: err(error),
          summaries: summaries,
        };
      }

      if (axios.isAxiosError(error)) {
        const message = JSON.stringify(error.response!.data);
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
        );
        if (error.response!.status >= 400 && error.response!.status < 500) {
          return {
            result: err(new UnhandledUserError(actionName, message, helpLink)),
            summaries: summaries,
          };
        } else {
          return {
            result: err(new UnhandledSystemError(actionName, message)),
            summaries: summaries,
          };
        }
      }

      const message = JSON.stringify(error);
      context.logProvider?.error(
        getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
      );
      return {
        result: err(new UnhandledSystemError(actionName, JSON.stringify(error))),
        summaries: summaries,
      };
    }
  }

  private validateArgs(args: UpdateAadAppArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.manifestTemplatePath !== "string" || !args.manifestTemplatePath) {
      invalidParameters.push("manifestTemplatePath");
    }

    if (typeof args.outputFilePath !== "string" || !args.outputFilePath) {
      invalidParameters.push("outputFilePath");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidParameterUserError(actionName, invalidParameters, helpLink);
    }
  }

  private loadCurrentState(): UpdateAadAppOutput {
    return {
      AAD_APP_ACCESS_AS_USER_PERMISSION_ID: process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID,
    };
  }

  private async loadManifest(
    manifestPath: string,
    state: UpdateAadAppOutput
  ): Promise<AADManifest> {
    let generatedNewPermissionId = false;
    try {
      const manifestTemplate = await fs.readFile(manifestPath, "utf8");
      const permissionIdPlaceholderRegex = /\${{ *AAD_APP_ACCESS_AS_USER_PERMISSION_ID *}}/;

      // generate a new permission id if there's no one in env and manifest needs it
      if (!process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID) {
        const matches = permissionIdPlaceholderRegex.exec(manifestTemplate);
        if (matches) {
          const permissionId = getUuid();
          process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID = permissionId;
          state.AAD_APP_ACCESS_AS_USER_PERMISSION_ID = permissionId;
          generatedNewPermissionId = true;
        }
      }

      const manifestString = expandEnvironmentVariable(manifestTemplate);
      this.validateManifestString(manifestString);
      const manifest: AADManifest = JSON.parse(manifestString);
      AadManifestHelper.processRequiredResourceAccessInManifest(manifest);
      return manifest;
    } finally {
      if (generatedNewPermissionId) {
        // restore environment variable to avoid impact to other code
        delete process.env.AAD_APP_ACCESS_AS_USER_PERMISSION_ID;
      }
    }
  }

  private getAbsolutePath(relativeOrAbsolutePath: string, projectPath: string) {
    return path.isAbsolute(relativeOrAbsolutePath)
      ? relativeOrAbsolutePath
      : path.join(projectPath, relativeOrAbsolutePath);
  }

  private getProgressBarSetting(): ProgressBarSetting {
    return {
      title: getLocalizedString("driver.aadApp.progressBar.updateAadAppTitle"),
      stepMessages: [
        getLocalizedString("driver.aadApp.progressBar.updateAadAppStepMessage"), // step 1
      ],
    };
  }

  private validateManifestString(manifestString: string) {
    const unresolvedEnvironmentVariable = getEnvironmentVariables(manifestString);
    if (unresolvedEnvironmentVariable && unresolvedEnvironmentVariable.length > 0) {
      throw new MissingEnvUserError(
        actionName,
        unresolvedEnvironmentVariable,
        helpLink,
        driverConstants.generateManifestFailedMessageKey
      );
    }
  }
}
