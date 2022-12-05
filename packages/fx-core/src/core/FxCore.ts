// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as os from "os";
import fs from "fs-extra";
import * as jsonschema from "jsonschema";
import * as path from "path";
import { Container } from "typedi";
import * as uuid from "uuid";
import { hooks } from "@feathersjs/hooks";
import {
  AppPackageFolderName,
  BuildFolderName,
  ConfigFolderName,
  CoreCallbackEvent,
  CoreCallbackFunc,
  DefaultReadme,
  err,
  Func,
  FunctionRouter,
  FxError,
  InputConfigsFolderName,
  Inputs,
  InputsWithProjectPath,
  ok,
  Platform,
  ProjectConfig,
  ProjectConfigV3,
  ProjectSettings,
  ProjectSettingsV3,
  QTreeNode,
  Result,
  Settings,
  Stage,
  StatesFolderName,
  Tools,
  UserCancelError,
  v2,
  v3,
  Void,
} from "@microsoft/teamsfx-api";

import { getLocalizedString } from "../common/localizeUtils";
import { localSettingsFileName } from "../common/localSettingsProvider";
import { isValidProject, newProjectSettings } from "../common/projectSettingsHelper";
import { TelemetryReporterInstance } from "../common/telemetry";
import { createV2Context, isV3Enabled } from "../common/tools";
import { getTemplatesFolder } from "../folder";
import {
  ApiConnectionOptionItem,
  AzureSolutionQuestionNames,
  CicdOptionItem,
  ExistingTabOptionItem,
  SingleSignOnOptionItem,
  ComponentNames,
  AadConstants,
} from "../component/constants";
import { CallbackRegistry } from "./callback";
import { checkPermission, grantPermission, listCollaborator } from "./collaborator";
import { LocalCrypto } from "./crypto";
import { environmentManager, newEnvInfoV3 } from "./environment";
import {
  CopyFileError,
  InvalidInputError,
  NotImplementedError,
  ObjectIsUndefinedError,
  OperationNotPermittedError,
  ProjectFolderExistError,
  TaskNotSupportError,
  WriteFileError,
  NoAadManifestExistError,
} from "./error";
import { setCurrentStage, setTools, TOOLS } from "./globalVars";
import { AadManifestMigrationMW } from "./middleware/aadManifestMigration";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ProjectConsolidateMW } from "./middleware/consolidateLocalRemote";
import { ContextInjectorMW } from "./middleware/contextInjector";
import { askNewEnvironment, EnvInfoLoaderMW_V3, loadEnvInfoV3 } from "./middleware/envInfoLoaderV3";
import { EnvInfoWriterMW_V3 } from "./middleware/envInfoWriterV3";
import { ErrorHandlerMW } from "./middleware/errorHandler";
import { ProjectMigratorMW } from "./middleware/projectMigrator";
import { ProjectSettingsLoaderMW } from "./middleware/projectSettingsLoader";
import { ProjectSettingsWriterMW } from "./middleware/projectSettingsWriter";
import { getQuestionsForCreateProjectV2, QuestionModelMW } from "./middleware/questionModel";
import { CoreQuestionNames, ProjectNamePattern } from "./question";
import {
  CoreTelemetryComponentName,
  CoreTelemetryEvent,
  CoreTelemetryProperty,
  CoreTelemetrySuccess,
  sendErrorTelemetryThenReturnError,
} from "./telemetry";
import { CoreHookContext } from "./types";
import { createContextV3, createDriverContext } from "../component/utils";
import { preCheck } from "../component/core";
import {
  FeatureId,
  getQuestionsForAddFeatureSubCommand,
  getQuestionsForAddFeatureV3,
  getQuestionsForAddResourceV3,
  getQuestionsForDeployV3,
  getQuestionsForInit,
  getQuestionsForProvisionV3,
} from "../component/question";
import { ProjectVersionCheckerMW } from "./middleware/projectVersionChecker";
import { addCicdQuestion } from "../component/feature/cicd/cicd";
import { AppManifest, publishQuestion } from "../component/resource/appManifest/appManifest";
import { ApiConnectorImpl } from "../component/feature/apiconnector/ApiConnectorImpl";
import { createEnvWithName } from "../component/envManager";
import { getProjectTemplatesFolderPath } from "../common/utils";
import { manifestUtils } from "../component/resource/appManifest/utils/ManifestUtils";
import { copyParameterJson } from "../component/arm";
import { ProjectSettingsHelper } from "../common/local";
import "../component/driver/aad/update";
import { UpdateAadAppArgs } from "../component/driver/aad/interface/updateAadAppArgs";
import { ValidateTeamsAppDriver } from "../component/driver/teamsApp/validate";
import { ValidateTeamsAppArgs } from "../component/driver/teamsApp/interfaces/ValidateTeamsAppArgs";
import { DriverContext } from "../component/driver/interface/commonArgs";
import { coordinator } from "../component/coordinator";
import { CreateAppPackageDriver } from "../component/driver/teamsApp/createAppPackage";
import { CreateAppPackageArgs } from "../component/driver/teamsApp/interfaces/CreateAppPackageArgs";
import { EnvLoaderMW, EnvWriterMW } from "../component/middleware/envMW";
import { envUtil } from "../component/utils/envUtil";
import { YamlParser } from "../component/configManager/parser";
import { ILifecycle, LifecycleName } from "../component/configManager/interface";
import "../component/driver/teamsApp/createAppPackage";
import "../component/driver/teamsApp/create";
import "../component/driver/teamsApp/configure";
import "../component/driver/teamsApp/copyAppPackageForSPFx";
import "../component/driver/teamsApp/publishAppPackage";
import "../component/driver/teamsApp/validate";
import "../component/driver/aad/create";
import "../component/driver/aad/update";
import "../component/driver/arm/deploy";
import "../component/driver/botAadApp/create";
import "../component/driver/deploy/azure/azureAppServiceDeployDriver";
import "../component/driver/deploy/azure/azureFunctionDeployDriver";
import "../component/driver/deploy/azure/azureStorageDeployDriver";
import "../component/driver/deploy/azure/azureStorageStaticWebsiteConfigDriver";
import "../component/driver/deploy/spfx/deployDriver";
import "../component/driver/script/dotnetBuildDriver";
import "../component/driver/script/npmBuildDriver";
import "../component/driver/script/npxBuildDriver";
import "../component/driver/tools/installDriver";
import "../component/driver/env/generate";
import "../component/driver/env/appsettingsGenerate";
import "../component/driver/botFramework/createOrUpdateBot";
import { settingsUtil } from "../component/utils/settingsUtil";
import { DotenvParseOutput } from "dotenv";
import {
  containsUnsupportedFeature,
  getFeaturesFromAppDefinition,
} from "../component/resource/appManifest/utils/utils";
import { VideoFilterAppBlockerMW } from "./middleware/videoFilterAppBlocker";
import { ProjectMigratorMWV3 } from "./middleware/projectMigratorV3";

export class FxCore implements v3.ICore {
  tools: Tools;
  isFromSample?: boolean;
  settingsVersion?: string;

  constructor(tools: Tools) {
    this.tools = tools;
    setTools(tools);
    TelemetryReporterInstance.telemetryReporter = tools.telemetryReporter;
  }

  /**
   * @todo this's a really primitive implement. Maybe could use Subscription Model to
   * refactor later.
   */
  public on(event: CoreCallbackEvent, callback: CoreCallbackFunc): void {
    return CallbackRegistry.set(event, callback);
  }

  async createExistingTabApp(
    inputs: Inputs,
    folder: string,
    ctx?: CoreHookContext
  ): Promise<Result<string, FxError>> {
    TOOLS.telemetryReporter?.sendTelemetryEvent(CoreTelemetryEvent.CreateStart, {
      [CoreTelemetryProperty.Component]: CoreTelemetryComponentName,
      [CoreTelemetryProperty.Capabilities]: ExistingTabOptionItem.id,
    });

    const appName = inputs[CoreQuestionNames.AppName] as string;
    inputs.folder = path.join(folder, appName);
    const result = await this._init(inputs, ctx, true);
    if (result.isErr()) {
      return err(
        sendErrorTelemetryThenReturnError(
          CoreTelemetryEvent.Create,
          result.error,
          TOOLS.telemetryReporter
        )
      );
    }

    TOOLS.ui.showMessage("info", getLocalizedString("core.create.successNotice"), false);
    TOOLS.telemetryReporter?.sendTelemetryEvent(CoreTelemetryEvent.Create, {
      [CoreTelemetryProperty.Component]: CoreTelemetryComponentName,
      [CoreTelemetryProperty.Success]: CoreTelemetrySuccess.Yes,
      [CoreTelemetryProperty.Capabilities]: ExistingTabOptionItem.id,
    });
    return result;
  }

  async createProject(inputs: Inputs): Promise<Result<string, FxError>> {
    if (isV3Enabled()) return this.createProjectNew(inputs);
    else return this.createProjectOld(inputs);
  }

  @hooks([ErrorHandlerMW, ContextInjectorMW])
  async createProjectNew(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    if (!ctx) {
      return err(new ObjectIsUndefinedError("ctx for createProject"));
    }
    setCurrentStage(Stage.create);
    inputs.stage = Stage.create;
    const context = createContextV3();
    if (!!inputs.teamsAppFromTdp) {
      // should never happen as we do same check on Developer Portal.
      if (containsUnsupportedFeature(inputs.teamsAppFromTdp)) {
        return err(InvalidInputError("Teams app contains unsupported features"));
      } else {
        context.telemetryReporter.sendTelemetryEvent(CoreTelemetryEvent.CreateFromTdpStart, {
          [CoreTelemetryProperty.TdpTeamsAppFeatures]: getFeaturesFromAppDefinition(
            inputs.teamsAppFromTdp
          ).join(","),
          [CoreTelemetryProperty.TdpTeamsAppId]: inputs.teamsAppFromTdp.teamsAppId,
        });
      }
    }
    const res = await coordinator.create(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    ctx.projectSettings = context.projectSetting;
    inputs.projectPath = context.projectPath;
    return ok(inputs.projectPath!);
  }

  /**
   * "teamsfx init infra" CLI command
   */
  @hooks([ErrorHandlerMW])
  async initInfra(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const res = await coordinator.initInfra(createContextV3(), inputs);
    return res;
  }
  /**
   * "teamsfx init debug" CLI command
   */
  @hooks([ErrorHandlerMW])
  async initDebug(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const res = await coordinator.initDebug(createContextV3(), inputs);
    return res;
  }
  @hooks([ErrorHandlerMW, ContextInjectorMW, ProjectSettingsWriterMW])
  async createProjectOld(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    if (!ctx) {
      return err(new ObjectIsUndefinedError("ctx for createProject"));
    }
    setCurrentStage(Stage.create);
    inputs.stage = Stage.create;
    const context = createContextV3();
    let res;
    if (isV3Enabled()) {
      res = await coordinator.create(context, inputs as InputsWithProjectPath);
    } else {
      const fx = Container.get("fx") as any;
      res = await fx.create(context, inputs as InputsWithProjectPath);
    }
    if (res.isErr()) return err(res.error);
    ctx.projectSettings = context.projectSetting;
    inputs.projectPath = context.projectPath;
    return ok(context.projectPath!);
  }

  async provisionResources(inputs: Inputs): Promise<Result<Void, FxError>> {
    return isV3Enabled() ? this.provisionResourcesNew(inputs) : this.provisionResourcesOld(inputs);
  }

  @hooks([ErrorHandlerMW, ProjectMigratorMWV3, EnvLoaderMW(false), ContextInjectorMW, EnvWriterMW])
  async provisionResourcesNew(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.provision);
    inputs.stage = Stage.provision;
    const context = createDriverContext(inputs);
    try {
      const [output, error] = await coordinator.provision(context, inputs as InputsWithProjectPath);
      ctx!.envVars = output;
      if (error) return err(error);
      return ok(Void);
    } finally {
      //reset subscription
      try {
        await TOOLS.tokenProvider.azureAccountProvider.setSubscription("");
      } catch (e) {}
    }
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    VideoFilterAppBlockerMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(),
  ])
  async provisionResourcesOld(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.provision);
    inputs.stage = Stage.provision;
    const context = createContextV3();
    context.envInfo = ctx!.envInfoV3!;
    context.projectSetting = ctx!.projectSettings! as ProjectSettingsV3;
    context.tokenProvider = TOOLS.tokenProvider;
    if (context.envInfo.envName === "local") {
      context.envInfo.config.isLocalDebug = true;
    }
    const fx = Container.get("fx") as any;
    const res = await fx.provision(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    ctx!.projectSettings = context.projectSetting;
    ctx!.envInfoV3 = context.envInfo;
    return ok(Void);
  }

  /**
   * Only used to provision Teams app with user provided app package
   * @param inputs
   * @returns teamsAppId on provision success
   */
  async provisionTeamsAppForCLI(inputs: Inputs): Promise<Result<string, FxError>> {
    if (!inputs.appPackagePath) {
      return err(InvalidInputError("appPackagePath is not defined", inputs));
    }
    const projectSettings: ProjectSettings = {
      appName: "fake",
      projectId: uuid.v4(),
    };
    const context: v2.Context = {
      userInteraction: TOOLS.ui,
      logProvider: TOOLS.logProvider,
      telemetryReporter: TOOLS.telemetryReporter!,
      cryptoProvider: new LocalCrypto(projectSettings.projectId),
      permissionRequestProvider: TOOLS.permissionRequest,
      projectSetting: projectSettings,
    };
    const appStudioV3 = Container.get<AppManifest>(ComponentNames.AppManifest);
    return appStudioV3.provisionForCLI(
      context,
      inputs as v2.InputsWithProjectPath,
      newEnvInfoV3(),
      TOOLS.tokenProvider
    );
  }

  async deployArtifacts(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    return isV3Enabled() ? this.deployArtifactsNew(inputs) : this.deployArtifactsOld(inputs);
  }

  @hooks([ErrorHandlerMW, ProjectMigratorMWV3, EnvLoaderMW(false), ContextInjectorMW, EnvWriterMW])
  async deployArtifactsNew(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.deploy);
    inputs.stage = Stage.deploy;
    const context = createDriverContext(inputs);
    const [output, error] = await coordinator.deploy(context, inputs as InputsWithProjectPath);
    ctx!.envVars = output;
    if (error) return err(error);
    return ok(Void);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    VideoFilterAppBlockerMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(),
  ])
  async deployArtifactsOld(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.deploy);
    inputs.stage = Stage.deploy;
    const context = createContextV3();
    context.envInfo = ctx!.envInfoV3!;
    context.projectSetting = ctx!.projectSettings! as ProjectSettingsV3;
    context.tokenProvider = TOOLS.tokenProvider;
    const fx = Container.get("fx") as any;
    const res = await fx.deploy(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    ctx!.projectSettings = context.projectSetting;
    return ok(Void);
  }
  async localDebug(inputs: Inputs): Promise<Result<Void, FxError>> {
    inputs.env = environmentManager.getLocalEnvName();
    return this.provisionResources(inputs);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
    EnvInfoWriterMW_V3(),
  ])
  async deployAadManifest(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.deployAad);
    inputs.stage = Stage.deployAad;
    const updateAadClient = Container.get("aadApp/update") as any;
    // In V3, the aad.template.json exist at .fx folder, and output to root build folder.
    const manifestTemplatePath: string = inputs.AAD_MANIFEST_FILE
      ? inputs.AAD_MANIFEST_FILE
      : path.join(inputs.projectPath!, AadConstants.DefaultTemplateFileName);
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new NoAadManifestExistError(manifestTemplatePath));
    }
    await fs.ensureDir(path.join(inputs.projectPath!, "build"));
    const manifestOutputPath: string = path.join(
      inputs.projectPath!,
      "build",
      `aad.${inputs.env}.json`
    );
    const inputArgs: UpdateAadAppArgs = {
      manifestTemplatePath: manifestTemplatePath,
      outputFilePath: manifestOutputPath,
    };
    const contextV3: DriverContext = {
      azureAccountProvider: TOOLS.tokenProvider.azureAccountProvider,
      m365TokenProvider: TOOLS.tokenProvider.m365TokenProvider,
      ui: TOOLS.ui,
      logProvider: TOOLS.logProvider,
      telemetryReporter: TOOLS.telemetryReporter!,
      projectPath: inputs.projectPath as string,
      platform: Platform.VSCode,
    };
    const res = await updateAadClient.run(inputArgs, contextV3);
    if (res.isErr()) return err(res.error);
    return ok(Void);
  }
  async publishApplication(inputs: Inputs): Promise<Result<Void, FxError>> {
    return isV3Enabled() ? this.publishApplicationNew(inputs) : this.publishApplicationOld(inputs);
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    VideoFilterAppBlockerMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(),
  ])
  async publishApplicationOld(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.publish);
    inputs.stage = Stage.publish;
    const context = createContextV3();
    context.envInfo = ctx!.envInfoV3!;
    context.projectSetting = ctx!.projectSettings! as ProjectSettingsV3;
    context.tokenProvider = TOOLS.tokenProvider;
    const appManifest = Container.get(ComponentNames.AppManifest) as any;
    const res = await appManifest.publish(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    ctx!.projectSettings = context.projectSetting;
    return ok(Void);
  }
  @hooks([ErrorHandlerMW, ProjectMigratorMWV3, EnvLoaderMW(false)])
  async publishApplicationNew(inputs: Inputs): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.publish);
    inputs.stage = Stage.publish;
    const context = createDriverContext(inputs);
    const res = await coordinator.publish(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    return ok(Void);
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(),
  ])
  async addFeature(
    inputs: v2.InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<any, FxError>> {
    inputs.stage = Stage.addFeature;
    const context = createContextV3(ctx?.projectSettings as ProjectSettingsV3);
    const fx = Container.get("fx") as any;
    const res = await fx.addFeature(context, inputs as InputsWithProjectPath);
    if (res.isErr()) return err(res.error);
    ctx!.projectSettings = context.projectSetting;
    return ok(res.value);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    VideoFilterAppBlockerMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
    ProjectSettingsWriterMW,
    EnvInfoWriterMW_V3(),
  ])
  async executeUserTask(
    func: Func,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<any, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("executeUserTask context"));
    let res: Result<any, FxError> = ok(undefined);
    const context = createContextV3(ctx?.projectSettings as ProjectSettingsV3);
    if (ctx?.envInfoV3) {
      context.envInfo = ctx.envInfoV3;
      if (context.envInfo.envName === "local") {
        context.envInfo.config.isLocalDebug = true;
      }
    }
    if (func.method === "addCICDWorkflows") {
      const component = Container.get("cicd") as any;
      inputs[AzureSolutionQuestionNames.Features] = CicdOptionItem.id;
      res = await component.add(context, inputs as InputsWithProjectPath);
    } else if (func.method === "connectExistingApi") {
      const component = Container.get("api-connector") as any;
      inputs[AzureSolutionQuestionNames.Features] = ApiConnectionOptionItem.id;
      res = await component.add(context, inputs as InputsWithProjectPath);
    } else if (func.method === "addSso") {
      inputs.stage = Stage.addFeature;
      inputs[AzureSolutionQuestionNames.Features] = SingleSignOnOptionItem.id;
      const component = Container.get("sso") as any;
      res = await component.add(context, inputs as InputsWithProjectPath);
    } else if (func.method === "addFeature") {
      inputs.stage = Stage.addFeature;
      const fx = Container.get("fx") as any;
      res = await fx.addFeature(context, inputs as InputsWithProjectPath);
    } else if (func.method === "getManifestTemplatePath") {
      const path = await manifestUtils.getTeamsAppManifestPath(
        (inputs as InputsWithProjectPath).projectPath
      );
      res = ok(path);
    } else if (func.method === "validateManifest") {
      if (isV3Enabled()) {
        const driver: ValidateTeamsAppDriver = Container.get("teamsApp/validate");
        const args: ValidateTeamsAppArgs = {
          manifestTemplatePath: func.params.manifestTemplatePath,
        };
        const driverContext: DriverContext = {
          azureAccountProvider: context.tokenProvider!.azureAccountProvider,
          m365TokenProvider: context.tokenProvider!.m365TokenProvider,
          ui: context.userInteraction,
          logProvider: context.logProvider,
          telemetryReporter: context.telemetryReporter,
          projectPath: inputs.projectPath!,
          platform: inputs.platform,
        };
        // await envUtil.readEnv(context.projectPath!, func.params.env);
        res = await driver.run(args, driverContext);
      } else {
        const component = Container.get("app-manifest") as any;
        res = await component.validate(context, inputs as InputsWithProjectPath);
      }
    } else if (func.method === "buildPackage") {
      if (isV3Enabled()) {
        const driver: CreateAppPackageDriver = Container.get("teamsApp/createAppPackage");
        const args: CreateAppPackageArgs = {
          manifestTemplatePath: func.params.manifestTemplatePath,
          outputZipPath: func.params.outputZipPath,
          outputJsonPath: func.params.outputJsonPath,
        };
        const driverContext: DriverContext = {
          azureAccountProvider: context.tokenProvider!.azureAccountProvider,
          m365TokenProvider: context.tokenProvider!.m365TokenProvider,
          ui: context.userInteraction,
          logProvider: context.logProvider,
          telemetryReporter: context.telemetryReporter,
          projectPath: inputs.projectPath!,
          platform: inputs.platform,
        };
        // await envUtil.readEnv(context.projectPath!, func.params.env);
        res = await driver.run(args, driverContext);
      } else {
        const component = Container.get("app-manifest") as any;
        res = await component.build(context, inputs as InputsWithProjectPath);
      }
    } else if (func.method === "updateManifest") {
      const component = Container.get("app-manifest") as any;
      res = await component.deploy(context, inputs as InputsWithProjectPath);
    } else if (func.method === "buildAadManifest") {
      const component = Container.get("aad-app") as any;
      res = await component.buildAadManifest(context, inputs as InputsWithProjectPath);
    } else {
      return err(new NotImplementedError(func.method));
    }
    if (res) {
      if (res.isErr()) return err(res.error);
      ctx!.projectSettings = context.projectSetting;
      return res;
    }
    return res;
  }

  @hooks([
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    ConcurrentLockerMW,
    EnvLoaderMW(true),
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployTeamsManifest(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    const context = createContextV3(ctx?.projectSettings as ProjectSettingsV3);
    const component = Container.get("app-manifest") as any;
    const res = await component.deployV3(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = envUtil.map2object(res.value);
    }
    return res;
  }

  /**
   * Warning: this API only works for CLI_HELP, it has no business with interactive run for CLI!
   */
  @hooks([ErrorHandlerMW])
  async getQuestions(
    stage: Stage,
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    inputs.stage = Stage.getQuestions;
    setCurrentStage(Stage.getQuestions);
    const context = createContextV3();
    if (stage === Stage.publish) {
      return await publishQuestion(inputs);
    } else if (stage === Stage.create) {
      return await getQuestionsForCreateProjectV2(inputs);
    } else if (stage === Stage.deploy) {
      return await getQuestionsForDeployV3(context, inputs);
    } else if (stage === Stage.provision) {
      return await getQuestionsForProvisionV3(context, inputs);
    } else if (stage === Stage.initDebug) {
      return await getQuestionsForInit("debug", inputs);
    } else if (stage === Stage.initInfra) {
      return await getQuestionsForInit("infra", inputs);
    }
    return ok(undefined);
  }

  async getQuestionsForAddFeature(
    featureId: FeatureId,
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const res = await getQuestionsForAddFeatureSubCommand(featureId, inputs);
    return res;
  }

  @hooks([ErrorHandlerMW])
  async getQuestionsForUserTask(
    func: FunctionRouter,
    inputs: Inputs
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    inputs.stage = Stage.getQuestions;
    setCurrentStage(Stage.getQuestions);
    const context = createContextV3();
    if (func.method === "addFeature") {
      return await getQuestionsForAddFeatureV3(context, inputs);
    } else if (func.method === "addResource") {
      return await getQuestionsForAddResourceV3(context, inputs);
    } else if (func.method === "addCICDWorkflows") {
      return await addCicdQuestion(context, inputs as InputsWithProjectPath);
    } else if (func.method === "connectExistingApi") {
      const apiConnectorImpl: ApiConnectorImpl = new ApiConnectorImpl();
      return await apiConnectorImpl.generateQuestion(context, inputs as InputsWithProjectPath);
    }
    return ok(undefined);
  }

  async getSettings(inputs: InputsWithProjectPath): Promise<Result<Settings, FxError>> {
    return settingsUtil.readSettings(inputs.projectPath);
  }

  @hooks([ErrorHandlerMW, EnvLoaderMW(true), ContextInjectorMW])
  async getDotEnv(
    inputs: InputsWithProjectPath,
    ctx?: CoreHookContext
  ): Promise<Result<DotenvParseOutput | undefined, FxError>> {
    return ok(ctx?.envVars);
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
  ])
  async getProjectConfig(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<ProjectConfig | undefined, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("getProjectConfig input stuff"));
    inputs.stage = Stage.getProjectConfig;
    setCurrentStage(Stage.getProjectConfig);
    return ok({
      settings: ctx.projectSettings,
      config: ctx.envInfoV3?.state,
    });
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    ContextInjectorMW,
  ])
  async getProjectConfigV3(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<ProjectConfigV3 | undefined, FxError>> {
    if (!ctx || !ctx.projectSettings)
      return err(new ObjectIsUndefinedError("getProjectConfigV3 input stuff"));
    if (!inputs.projectPath) return ok(undefined);
    inputs.stage = Stage.getProjectConfig;
    setCurrentStage(Stage.getProjectConfig);
    const config: ProjectConfigV3 = {
      projectSettings: ctx.projectSettings,
      envInfos: {},
    };
    const envNamesRes = await environmentManager.listAllEnvConfigs(inputs.projectPath);
    if (envNamesRes.isErr()) {
      return err(envNamesRes.error);
    }
    for (const env of envNamesRes.value) {
      const result = await loadEnvInfoV3(
        inputs as v2.InputsWithProjectPath,
        ctx.projectSettings,
        env,
        false
      );
      if (result.isErr()) {
        return err(result.error);
      }
      config.envInfos[env] = result.value;
    }
    return ok(config);
  }
  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW_V3(false),
    QuestionModelMW,
    ContextInjectorMW,
  ])
  async grantPermission(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    setCurrentStage(Stage.grantPermission);
    inputs.stage = Stage.grantPermission;
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(new ObjectIsUndefinedError("projectPath"));
    }
    if (ctx && ctx.contextV2 && (isV3Enabled() || ctx.envInfoV3)) {
      const context = createContextV3(ctx?.projectSettings as ProjectSettingsV3);
      context.envInfo = ctx.envInfoV3;
      const res = await grantPermission(
        context,
        inputs as v2.InputsWithProjectPath,
        ctx.envInfoV3,
        TOOLS.tokenProvider
      );
      return res;
    }
    return err(new ObjectIsUndefinedError("ctx, contextV2, envInfoV3"));
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
  ])
  async checkPermission(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    setCurrentStage(Stage.checkPermission);
    inputs.stage = Stage.checkPermission;
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(new ObjectIsUndefinedError("projectPath"));
    }
    if (ctx && ctx.contextV2 && (isV3Enabled() || ctx.envInfoV3)) {
      const context = createContextV3(ctx?.projectSettings as ProjectSettingsV3);
      context.envInfo = ctx.envInfoV3;
      const res = await checkPermission(
        context,
        inputs as v2.InputsWithProjectPath,
        ctx.envInfoV3,
        TOOLS.tokenProvider
      );
      return res;
    }
    return err(new ObjectIsUndefinedError("ctx, contextV2, envInfoV3"));
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectMigratorMW,
    ProjectConsolidateMW,
    AadManifestMigrationMW,
    ProjectVersionCheckerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
  ])
  async listCollaborator(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
    setCurrentStage(Stage.listCollaborator);
    inputs.stage = Stage.listCollaborator;
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(new ObjectIsUndefinedError("projectPath"));
    }
    if (ctx && ctx.contextV2 && (isV3Enabled() || ctx.envInfoV3)) {
      const context = createContextV3(ctx?.projectSettings as ProjectSettingsV3);
      context.envInfo = ctx.envInfoV3;
      const res = await listCollaborator(
        context,
        inputs as v2.InputsWithProjectPath,
        ctx.envInfoV3,
        TOOLS.tokenProvider
      );
      return res;
    }
    return err(new ObjectIsUndefinedError("ctx, contextV2, envInfoV3"));
  }

  @hooks([
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ProjectSettingsLoaderMW,
    EnvInfoLoaderMW_V3(false),
    ContextInjectorMW,
  ])
  async getSelectedEnv(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string | undefined, FxError>> {
    return ok(ctx?.envInfoV3?.envName);
  }

  @hooks([ErrorHandlerMW, ConcurrentLockerMW, ProjectSettingsLoaderMW, ContextInjectorMW])
  async encrypt(
    plaintext: string,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("ctx"));
    if (!ctx.contextV2) return err(new ObjectIsUndefinedError("ctx.contextV2"));
    return ctx.contextV2.cryptoProvider.encrypt(plaintext);
  }

  @hooks([ErrorHandlerMW, ConcurrentLockerMW, ProjectSettingsLoaderMW, ContextInjectorMW])
  async decrypt(
    ciphertext: string,
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<string, FxError>> {
    if (!ctx) return err(new ObjectIsUndefinedError("ctx"));
    if (!ctx.contextV2) return err(new ObjectIsUndefinedError("ctx.contextV2"));
    return ctx.contextV2.cryptoProvider.decrypt(ciphertext);
  }

  async buildArtifacts(inputs: Inputs): Promise<Result<Void, FxError>> {
    throw new TaskNotSupportError(Stage.build);
  }

  @hooks([ErrorHandlerMW, ConcurrentLockerMW, ProjectSettingsLoaderMW, ContextInjectorMW])
  async createEnv(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    if (!ctx || !inputs.projectPath)
      return err(new ObjectIsUndefinedError("createEnv input stuff"));
    const projectSettings = ctx.projectSettings;
    if (!projectSettings) {
      return ok(Void);
    }

    const core = ctx!.self as FxCore;
    const createEnvCopyInput = await askNewEnvironment(ctx!, inputs);

    if (
      !createEnvCopyInput ||
      !createEnvCopyInput.targetEnvName ||
      !createEnvCopyInput.sourceEnvName
    ) {
      return err(UserCancelError);
    }

    if (isV3Enabled()) {
      return this.createEnvCopyV3(
        createEnvCopyInput.targetEnvName,
        createEnvCopyInput.sourceEnvName,
        inputs.projectPath
      );
    }

    const createEnvResult = await this.createEnvCopy(
      createEnvCopyInput.targetEnvName,
      createEnvCopyInput.sourceEnvName,
      inputs,
      core
    );

    if (createEnvResult.isErr()) {
      return createEnvResult;
    }

    inputs.sourceEnvName = createEnvCopyInput.sourceEnvName;
    inputs.targetEnvName = createEnvCopyInput.targetEnvName;

    if (!ProjectSettingsHelper.isSpfx(ctx.projectSettings)) {
      await copyParameterJson(
        inputs.projectPath!,
        ctx.projectSettings!.appName!,
        inputs.targetEnvName!,
        inputs.sourceEnvName!
      );
    }

    return ok(Void);
  }

  // a phantom migration method for V3
  @hooks([ErrorHandlerMW, ProjectMigratorMWV3])
  async phantomMigrationV3(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    return ok(Void);
  }

  async createEnvCopyV3(
    targetEnvName: string,
    sourceEnvName: string,
    projectPath: string
  ): Promise<Result<Void, FxError>> {
    const sourceDotEnvFile = environmentManager.getDotEnvPath(sourceEnvName, projectPath);
    const source = await fs.readFile(sourceDotEnvFile);
    const targetDotEnvFile = environmentManager.getDotEnvPath(targetEnvName, projectPath);
    const writeStream = fs.createWriteStream(targetDotEnvFile);
    source
      .toString()
      .split(/\r?\n/)
      .forEach((line) => {
        const reg = /^([a-zA-Z_][a-zA-Z0-9_]*=)/g;
        const match = reg.exec(line);
        if (match) {
          if (match[1].startsWith("TEAMSFX_ENV=")) {
            writeStream.write(`TEAMSFX_ENV=${targetEnvName}${os.EOL}`);
          } else {
            writeStream.write(`${match[1]}${os.EOL}`);
          }
        } else {
          writeStream.write(`${line.trim()}${os.EOL}`);
        }
      });

    writeStream.end();
    return ok(Void);
  }

  async createEnvCopy(
    targetEnvName: string,
    sourceEnvName: string,
    inputs: Inputs,
    core: FxCore
  ): Promise<Result<Void, FxError>> {
    // copy env config file
    const targetEnvConfigFilePath = environmentManager.getEnvConfigPath(
      targetEnvName,
      inputs.projectPath!
    );
    const sourceEnvConfigFilePath = environmentManager.getEnvConfigPath(
      sourceEnvName,
      inputs.projectPath!
    );

    try {
      await fs.copy(sourceEnvConfigFilePath, targetEnvConfigFilePath);
    } catch (e) {
      return err(CopyFileError(e as Error));
    }

    TOOLS.logProvider.debug(
      `[core] copy env config file for ${targetEnvName} environment to path ${targetEnvConfigFilePath}`
    );

    return ok(Void);
  }

  async activateEnv(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<Void, FxError>> {
    return ok(Void);
  }

  // apply the given yaml template to current project.
  async apply(
    inputs: Inputs,
    templatePath: string,
    lifecycleName: string
  ): Promise<Result<Void, FxError>> {
    if (!inputs.projectPath) {
      return err(InvalidInputError("invalid projectPath", inputs));
    }
    const projectPath = inputs.projectPath;
    if (!inputs.env) {
      return err(InvalidInputError("invalid env", inputs));
    }
    const env = inputs.env;
    const lifecycleName_: LifecycleName = lifecycleName as LifecycleName;
    const result = await envUtil.readEnv(projectPath, env);
    if (result.isErr()) {
      return err(result.error);
    }

    const parser = new YamlParser();
    const maybeProjectModel = await parser.parse(templatePath);
    if (maybeProjectModel.isErr()) {
      return err(maybeProjectModel.error);
    }

    const projectModel = maybeProjectModel.value;
    const driverContext: DriverContext = {
      azureAccountProvider: TOOLS.tokenProvider.azureAccountProvider!,
      m365TokenProvider: TOOLS.tokenProvider.m365TokenProvider!,
      ui: TOOLS.ui,
      logProvider: TOOLS.logProvider,
      telemetryReporter: TOOLS.telemetryReporter!,
      projectPath: projectPath,
      platform: inputs.platform,
    };
    const lifecycle = projectModel[lifecycleName_];
    if (lifecycle) {
      return this.runLifecycle(lifecycle, driverContext, env);
    } else {
      await driverContext.logProvider.warning(`No definition found for ${lifecycleName}`);
      return ok(Void);
    }
  }

  async runLifecycle(
    lifecycle: ILifecycle,
    driverContext: DriverContext,
    env: string
  ): Promise<Result<Void, FxError>> {
    const r = await lifecycle.execute(driverContext);
    const runResult = r.result;
    if (runResult.isOk()) {
      await driverContext.logProvider.info(`Lifecycle ${lifecycle.name} succeeded`);
      const writeResult = await envUtil.writeEnv(
        driverContext.projectPath,
        env,
        envUtil.map2object(runResult.value)
      );
      return writeResult.map(() => Void);
    } else {
      const error = runResult.error;
      if (error.kind === "Failure") {
        await driverContext.logProvider.error(
          `Failed to run ${lifecycle.name} due to ${error.error.name}: ${error.error.message}`
        );
        return err(error.error);
      } else {
        try {
          const failedDriver = error.reason.failedDriver;
          if (error.reason.kind === "UnresolvedPlaceholders") {
            const unresolved = error.reason.unresolvedPlaceHolders;
            await driverContext.logProvider.warning(
              `Unresolved placeholders: ${unresolved.join(",")} for driver ${failedDriver.uses}`
            );
            return ok(Void);
          } else {
            await driverContext.logProvider.error(
              `Failed to run ${lifecycle.name} due to ${error.reason.error.name}: ${error.reason.error.message}. Failed driver: ${failedDriver.uses}`
            );
            return err(error.reason.error);
          }
        } finally {
          await envUtil.writeEnv(driverContext.projectPath, env, envUtil.map2object(error.env));
        }
      }
    }
  }

  async _init(
    inputs: Inputs,
    ctx?: CoreHookContext,
    isInitExistingApp = false
  ): Promise<Result<string, FxError>> {
    if (!ctx) {
      return err(new ObjectIsUndefinedError("ctx for createProject"));
    }
    // validate app name
    const appName = inputs[CoreQuestionNames.AppName] as string;
    const validateResult = jsonschema.validate(appName, {
      pattern: ProjectNamePattern,
    });
    if (validateResult.errors && validateResult.errors.length > 0) {
      return err(InvalidInputError("invalid app-name", inputs));
    }

    const projectPath = inputs.folder;
    if (!projectPath) {
      return err(InvalidInputError("projectPath is empty", inputs));
    }

    if (isInitExistingApp) {
      const folderExist = await fs.pathExists(projectPath);
      if (folderExist) {
        return err(new ProjectFolderExistError(projectPath));
      }
    } else {
      const isValid = isValidProject(projectPath);
      if (isValid) {
        return err(
          new OperationNotPermittedError("initialize a project in existing teamsfx project")
        );
      }
    }

    await fs.ensureDir(projectPath);
    inputs.projectPath = projectPath;

    // create ProjectSettings
    const projectSettings = newProjectSettings();
    projectSettings.appName = appName;
    (projectSettings as ProjectSettingsV3).components = [];
    ctx.projectSettings = projectSettings;

    // create folder structure
    await fs.ensureDir(path.join(projectPath, `.${ConfigFolderName}`));
    await fs.ensureDir(
      path.join(await getProjectTemplatesFolderPath(projectPath), `${AppPackageFolderName}`)
    );
    const basicFolderRes = await ensureBasicFolderStructure(inputs, false);
    if (basicFolderRes.isErr()) {
      return err(basicFolderRes.error);
    }

    // create contextV2
    const context = createV2Context(projectSettings);
    ctx.contextV2 = context;

    const appStudioComponent = Container.get<AppManifest>(ComponentNames.AppManifest);

    // pre-check before initialize
    const preCheckResult = await preCheck(projectPath);
    if (preCheckResult.isErr()) {
      return err(preCheckResult.error);
    }

    // init manifest
    const manifestInitRes = await appStudioComponent.init(
      context,
      inputs as v2.InputsWithProjectPath,
      isInitExistingApp
    );
    if (manifestInitRes.isErr()) return err(manifestInitRes.error);

    const manifestAddcapRes = await appStudioComponent.addCapability(
      inputs as v2.InputsWithProjectPath,
      [{ name: "staticTab", existingApp: true }]
    );
    if (manifestAddcapRes.isErr()) return err(manifestAddcapRes.error);

    // create env config with existing tab's endpoint
    const endpoint = inputs[CoreQuestionNames.ExistingTabEndpoint] as string;
    const createEnvResult = await createEnvWithName(
      environmentManager.getDefaultEnvName(),
      projectSettings.appName,
      inputs as InputsWithProjectPath,
      isInitExistingApp ? endpoint : undefined
    );
    if (createEnvResult.isErr()) {
      return err(createEnvResult.error);
    }
    const createLocalEnvResult = await createEnvWithName(
      environmentManager.getLocalEnvName(),
      projectSettings.appName,
      inputs as InputsWithProjectPath,
      isInitExistingApp ? endpoint : undefined
    );
    if (createLocalEnvResult.isErr()) {
      return err(createLocalEnvResult.error);
    }
    const sourceReadmePath = path.join(getTemplatesFolder(), "core", DefaultReadme);
    if (await fs.pathExists(sourceReadmePath)) {
      const targetReadmePath = path.join(projectPath, DefaultReadme);
      await fs.copy(sourceReadmePath, targetReadmePath);
    }
    return ok(inputs.projectPath!);
  }

  @hooks([ErrorHandlerMW, ContextInjectorMW, ProjectSettingsWriterMW])
  async init(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<string, FxError>> {
    const result = await this._init(inputs, ctx);
    if (result.isOk()) {
      TOOLS.ui.showMessage("info", getLocalizedString("core.init.successNotice"), false);
    }

    return result;
  }

  @hooks([ErrorHandlerMW, EnvLoaderMW(false), ContextInjectorMW])
  async preProvisionForVS(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<
    Result<
      {
        needAzureLogin: boolean;
        needM365Login: boolean;
        resolvedAzureSubscriptionId?: string;
        resolvedAzureResourceGroupName?: string;
      },
      FxError
    >
  > {
    const context = createDriverContext(inputs);
    return coordinator.preProvisionForVS(context, inputs as InputsWithProjectPath);
  }

  @hooks([ErrorHandlerMW, EnvLoaderMW(false), ContextInjectorMW])
  async publishInDeveloperPortal(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<Void, FxError>> {
    setCurrentStage(Stage.publishInDeveloperPortal);
    inputs.stage = Stage.publishInDeveloperPortal;
    const context = createContextV3();
    return await coordinator.publishInDeveloperPortal(context, inputs as InputsWithProjectPath);
  }
}

export async function ensureBasicFolderStructure(
  inputs: Inputs,
  createPackageJson = true
): Promise<Result<null, FxError>> {
  if (!inputs.projectPath) {
    return err(new ObjectIsUndefinedError("projectPath"));
  }
  try {
    if (createPackageJson) {
      const appName = inputs[CoreQuestionNames.AppName] as string;
      if (inputs.platform !== Platform.VS) {
        const packageJsonFilePath = path.join(inputs.projectPath, `package.json`);
        const exists = await fs.pathExists(packageJsonFilePath);
        if (!exists) {
          await fs.writeFile(
            packageJsonFilePath,
            JSON.stringify(
              {
                name: appName,
                version: "0.0.1",
                description: "",
                author: "",
                scripts: {
                  test: 'echo "Error: no test specified" && exit 1',
                },
                devDependencies: {
                  "@microsoft/teamsfx-cli": "1.*",
                },
                license: "MIT",
              },
              null,
              4
            )
          );
        }
      }
    }
    {
      const gitIgnoreFilePath = path.join(inputs.projectPath, `.gitignore`);
      let lines: string[] = [];
      const exists = await fs.pathExists(gitIgnoreFilePath);
      if (exists) {
        const content = await fs.readFile(gitIgnoreFilePath, { encoding: "utf8" });
        lines = content.split("\n");
        for (let i = 0; i < lines.length; ++i) {
          lines[i] = lines[i].trim();
        }
      }
      const gitIgnoreContent = [
        "\n# TeamsFx files",
        "node_modules",
        `.${ConfigFolderName}/${InputConfigsFolderName}/${localSettingsFileName}`,
        `.${ConfigFolderName}/${StatesFolderName}/*.userdata`,
        ".DS_Store",
        ".env.teamsfx.local",
        "subscriptionInfo.json",
        BuildFolderName,
      ];
      gitIgnoreContent.push(`.${ConfigFolderName}/${InputConfigsFolderName}/config.local.json`);
      gitIgnoreContent.push(`.${ConfigFolderName}/${StatesFolderName}/state.local.json`);
      if (inputs.platform === Platform.VS) {
        gitIgnoreContent.push("appsettings.Development.json");
      }
      gitIgnoreContent.forEach((line) => {
        if (!lines.includes(line.trim())) {
          lines.push(line.trim());
        }
      });
      await fs.writeFile(gitIgnoreFilePath, lines.join("\n"), { encoding: "utf8" });
    }
  } catch (e) {
    return err(WriteFileError(e));
  }
  return ok(null);
}
