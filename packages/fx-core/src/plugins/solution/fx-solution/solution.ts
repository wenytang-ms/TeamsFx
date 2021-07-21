/* eslint-disable @typescript-eslint/ban-types */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ok,
  err,
  QTreeNode,
  ConfigMap,
  Stage,
  returnSystemError,
  returnUserError,
  PluginContext,
  Plugin,
  Func,
  FxError,
  Result,
  Solution,
  SolutionConfig,
  SolutionContext,
  TeamsAppManifest,
  OptionItem,
  ConfigFolderName,
  AzureSolutionSettings,
  Platform,
  Inputs,
  DynamicPlatforms,
  SubscriptionInfo,
} from "@microsoft/teamsfx-api";
import { checkSubscription, fillInCommonQuestions } from "./commonQuestions";
import { executeLifecycles, executeConcurrently, LifecyclesWithContext } from "./executor";
import {
  getPluginContext,
  getSubsriptionDisplayName,
  sendErrorTelemetryThenReturnError,
} from "./util";
import * as fs from "fs-extra";
import {
  DEFAULT_PERMISSION_REQUEST,
  GLOBAL_CONFIG,
  PERMISSION_REQUEST,
  SolutionError,
  LOCAL_DEBUG_AAD_ID,
  LOCAL_DEBUG_TEAMS_APP_ID,
  Void,
  SOLUTION_PROVISION_SUCCEEDED,
  LOCAL_APPLICATION_ID_URIS,
  LOCAL_CLIENT_SECRET,
  REMOTE_AAD_ID,
  REMOTE_APPLICATION_ID_URIS,
  REMOTE_CLIENT_SECRET,
  PROGRAMMING_LANGUAGE,
  REMOTE_MANIFEST,
  REMOTE_TEAMS_APP_ID,
  CancelError,
  SolutionTelemetryProperty,
  SolutionTelemetryEvent,
  SolutionTelemetryComponentName,
  SolutionTelemetrySuccess,
  PluginNames,
} from "./constants";

import {
  AzureResourceFunction,
  AzureResourceSQL,
  AzureResourcesQuestion,
  AzureSolutionQuestionNames,
  BotOptionItem,
  DeployPluginSelectQuestion,
  HostTypeOptionAzure,
  HostTypeOptionSPFx,
  FrontendHostTypeQuestion,
  TabOptionItem,
  MessageExtensionItem,
  AzureResourceApim,
  createCapabilityQuestion,
  createAddAzureResourceQuestion,
  AskSubscriptionQuestion,
  addCapabilityQuestion,
  ProgrammingLanguageQuestion,
} from "./question";
import Mustache from "mustache";
import path from "path";
import * as util from "util";
import { deepCopy, getStrings, isUserCancelError } from "../../../common/tools";
import { getTemplatesFolder } from "../../..";
import {
  getActivatedResourcePlugins,
  getAllResourcePluginMap,
  getAllResourcePlugins,
  ResourcePlugins,
} from "./ResourcePluginContainer";
import { AadAppForTeamsPlugin, AppStudioPlugin, SpfxPlugin } from "../../resource";
import { ErrorHandlerMW } from "../../../core/middleware/errorHandler";
import { hooks } from "@feathersjs/hooks/lib";
import { Service, Container } from "typedi";

export type LoadedPlugin = Plugin;
export type PluginsWithContext = [LoadedPlugin, PluginContext];

type ParamForRegisterTeamsAppAndAad = {
  "app-name": string;
  environment: "local" | "remote";
  endpoint: string;
  "root-path": string;
};

// Maybe we need a state machine to track state transition.
export enum SolutionRunningState {
  Idle = "idle",
  ProvisionInProgress = "ProvisionInProgress",
  DeployInProgress = "DeployInProgress",
  PublishInProgress = "PublishInProgress",
}

@Service()
export class TeamsAppSolution implements Solution {
  SpfxPlugin: SpfxPlugin;
  AppStudioPlugin: AppStudioPlugin;
  BotPlugin: Plugin;
  AadPlugin: Plugin;
  FrontendPlugin: Plugin;
  FunctionPlugin: Plugin;
  SqlPlugin: Plugin;
  ApimPlugin: Plugin;
  LocalDebugPlugin: Plugin;

  name = "fx-solution-azure";

  runningState: SolutionRunningState;

  constructor() {
    this.SpfxPlugin = Container.get<SpfxPlugin>(ResourcePlugins.SpfxPlugin);
    this.AppStudioPlugin = Container.get<AppStudioPlugin>(ResourcePlugins.AppStudioPlugin);
    this.BotPlugin = Container.get<Plugin>(ResourcePlugins.BotPlugin);
    this.AadPlugin = Container.get<Plugin>(ResourcePlugins.AadPlugin);
    this.FrontendPlugin = Container.get<Plugin>(ResourcePlugins.FrontendPlugin);
    this.FunctionPlugin = Container.get<Plugin>(ResourcePlugins.FunctionPlugin);
    this.SqlPlugin = Container.get<Plugin>(ResourcePlugins.SqlPlugin);
    this.ApimPlugin = Container.get<Plugin>(ResourcePlugins.ApimPlugin);
    this.LocalDebugPlugin = Container.get<Plugin>(ResourcePlugins.LocalDebugPlugin);
    this.runningState = SolutionRunningState.Idle;
  }

  private getPluginAndContextArray(
    ctx: SolutionContext,
    selectedPlugins: LoadedPlugin[],
    manifest: TeamsAppManifest
  ): PluginsWithContext[] {
    // let pluginContextConstructor = getPluginContextConstructor(ctx);
    return selectedPlugins.map((plugin) => [plugin, getPluginContext(ctx, plugin.name, manifest)]);
  }

  async init(ctx: SolutionContext): Promise<Result<any, FxError>> {
    return ok({});
  }

  fillInSolutionSettings(ctx: SolutionContext): Result<AzureSolutionSettings, FxError> {
    const answers = ctx.answers;
    if (!answers) {
      return err(
        returnSystemError(new Error("answer is undefined"), "Solution", SolutionError.InternelError)
      );
    }
    const projectSettings = ctx.projectSettings;
    if (!projectSettings) {
      return err(
        returnSystemError(
          new Error("projectSettings is undefined"),
          "Solution",
          SolutionError.InternelError
        )
      );
    }
    if (!projectSettings.solutionSettings) {
      return err(
        returnSystemError(
          new Error("solutionSettings is undefined"),
          "Solution",
          SolutionError.InternelError
        )
      );
    }
    const capabilities = (answers[AzureSolutionQuestionNames.Capabilities] as string[]) || [];
    if (!capabilities || capabilities.length === 0) {
      return err(
        returnSystemError(
          new Error("capabilities is empty"),
          "Solution",
          SolutionError.InternelError
        )
      );
    }
    let hostType = answers[AzureSolutionQuestionNames.HostType] as string;
    if (capabilities.includes(BotOptionItem.id) || capabilities.includes(MessageExtensionItem.id))
      hostType = HostTypeOptionAzure.id;
    if (!hostType) {
      return err(
        returnSystemError(
          new Error("hostType is undefined"),
          "Solution",
          SolutionError.InternelError
        )
      );
    }
    let azureResources: string[] | undefined;
    if (hostType === HostTypeOptionAzure.id && capabilities.includes(TabOptionItem.id)) {
      azureResources = answers[AzureSolutionQuestionNames.AzureResources] as string[];
      if (azureResources) {
        if (
          (azureResources.includes(AzureResourceSQL.id) ||
            azureResources.includes(AzureResourceApim.id)) &&
          !azureResources.includes(AzureResourceFunction.id)
        ) {
          azureResources.push(AzureResourceFunction.id);
        }
      } else azureResources = [];
    }
    const solutionSettings: AzureSolutionSettings = {
      name: projectSettings.solutionSettings.name,
      version: projectSettings.solutionSettings.version,
      hostType: hostType,
      capabilities: capabilities,
      azureResources: azureResources || [],
      activeResourcePlugins: [],
    };
    projectSettings.solutionSettings = solutionSettings;
    return ok(solutionSettings);
  }

  /**
   * create
   */
  async create(ctx: SolutionContext): Promise<Result<any, FxError>> {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.CreateStart, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
    });

    // ensure that global namespace is present
    if (!ctx.config.has(GLOBAL_CONFIG)) {
      ctx.config.set(GLOBAL_CONFIG, new ConfigMap());
    }

    // Only non-SPFx project will ask this question.
    const lang = ctx.answers![AzureSolutionQuestionNames.ProgrammingLanguage] as string;
    if (lang) {
      ctx.config.get(GLOBAL_CONFIG)?.set(PROGRAMMING_LANGUAGE, lang);
    }

    const settingsRes = this.fillInSolutionSettings(ctx);
    if (settingsRes.isErr()) {
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.Create,
          settingsRes.error,
          ctx.telemetryReporter
        )
      );
    }

    const solutionSettings = settingsRes.value;

    //Reload plugins according to user answers
    await this.reloadPlugins(solutionSettings);

    const templatesFolder = getTemplatesFolder();
    const defaultColorPath = path.join(templatesFolder, "plugins", "solution", "defaultIcon.png");
    const defaultOutlinePath = path.join(
      templatesFolder,
      "plugins",
      "solution",
      "defaultOutline.png"
    );

    await fs.copy(defaultColorPath, `${ctx.root}/.${ConfigFolderName}/color.png`);
    await fs.copy(defaultOutlinePath, `${ctx.root}/.${ConfigFolderName}/outline.png`);
    if (this.isAzureProject(ctx)) {
      const manifest = await this.AppStudioPlugin.createManifest(ctx.projectSettings!);
      // if (manifest) Object.assign(ctx.app, manifest);
      await fs.writeFile(
        `${ctx.root}/.${ConfigFolderName}/${REMOTE_MANIFEST}`,
        JSON.stringify(manifest, null, 4)
      );
      await fs.writeJSON(`${ctx.root}/permissions.json`, DEFAULT_PERMISSION_REQUEST, { spaces: 4 });
      ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.Create, {
        [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
        [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
        [SolutionTelemetryProperty.Resources]: solutionSettings.azureResources.join(";"),
        [SolutionTelemetryProperty.Capabilities]: solutionSettings.capabilities.join(";"),
      });
    } else {
      const manifest = await (this.SpfxPlugin as SpfxPlugin).getManifest();
      await fs.writeFile(
        `${ctx.root}/.${ConfigFolderName}/${REMOTE_MANIFEST}`,
        JSON.stringify(manifest, null, 4)
      );
    }
    return ok(Void);
  }

  reloadPlugins(solutionSettings: AzureSolutionSettings) {
    const res = getActivatedResourcePlugins(solutionSettings);
    solutionSettings.activeResourcePlugins = res.map((p) => p.name);
  }

  private spfxSelected(ctx: SolutionContext): boolean {
    // Generally, if SPFx is selected, there should be no other plugins. But we don't check this invariant here.
    const spfxExists = this.getAzureSolutionSettings(ctx).activeResourcePlugins.some(
      (pluginName) => pluginName === this.SpfxPlugin.name
    );
    return spfxExists === undefined ? false : spfxExists;
  }

  private isAzureProject(ctx?: SolutionContext): boolean {
    if (!ctx) return true;
    const settings = this.getAzureSolutionSettings(ctx);
    return HostTypeOptionAzure.id === settings.hostType;
  }

  async update(ctx: SolutionContext): Promise<Result<any, FxError>> {
    return await this.executeAddResource(ctx);
  }

  private getSelectedPlugins(ctx: SolutionContext): Result<Plugin[], FxError> {
    const settings = this.getAzureSolutionSettings(ctx);
    const map = getAllResourcePluginMap();
    const results: Plugin[] = [];
    for (const name of settings.activeResourcePlugins) {
      const plugin = map.get(name);
      if (!plugin) {
        return err(
          returnUserError(
            new Error(`Plugin name ${name} is not valid`),
            "Solution",
            SolutionError.PluginNotFound
          )
        );
      }
      results.push(plugin);
    }
    return ok(results);
  }

  /**
   * scaffold
   */
  @hooks([ErrorHandlerMW])
  async scaffold(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const maybeSelectedPlugins = this.getSelectedPlugins(ctx);
    if (maybeSelectedPlugins.isErr()) {
      return maybeSelectedPlugins;
    }
    const selectedPlugins = maybeSelectedPlugins.value;
    const result = await this.doScaffold(ctx, selectedPlugins);
    if (result.isOk()) {
      ctx.ui?.showMessage("info", getStrings().solution.ScaffoldSuccessNotice, false);
    }
    return result;
  }

  async doScaffold(
    ctx: SolutionContext,
    selectedPlugins: LoadedPlugin[]
  ): Promise<Result<any, FxError>> {
    const appStudioPlugin = this.AppStudioPlugin as AppStudioPlugin;
    const maybeManifest = await appStudioPlugin.reloadManifestAndCheckRequiredFields(ctx.root);
    if (maybeManifest.isErr()) {
      return maybeManifest;
    }
    const manifest = maybeManifest.value;

    const pluginsWithCtx: PluginsWithContext[] = this.getPluginAndContextArray(
      ctx,
      selectedPlugins,
      manifest
    );
    const preScaffoldWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.preScaffold?.bind(plugin), context, plugin.name];
    });
    const scaffoldWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.scaffold?.bind(plugin), context, plugin.name];
    });
    const postScaffoldWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.postScaffold?.bind(plugin), context, plugin.name];
    });

    const res = await executeLifecycles(preScaffoldWithCtx, scaffoldWithCtx, postScaffoldWithCtx);

    if (res.isOk()) {
      const capabilities = (ctx.projectSettings?.solutionSettings as AzureSolutionSettings)
        .capabilities;
      const hasBot = capabilities?.includes(BotOptionItem.id);
      const hasMsgExt = capabilities?.includes(MessageExtensionItem.id);
      const hasTab = capabilities?.includes(TabOptionItem.id);
      if (hasTab && (hasBot || hasMsgExt)) {
        const readme = path.join(getTemplatesFolder(), "plugins", "solution", "README.md");
        if (await fs.pathExists(readme)) {
          await fs.copy(readme, `${ctx.root}/README.md`);
        }
      }
    }

    return res;
  }

  /**
   * Load the content of the latest permissions.json file to config
   * @param rootPath root path of this project
   * @param config solution config
   */
  private async getPermissionRequest(ctx: SolutionContext): Promise<Result<string, FxError>> {
    if (!this.isAzureProject(ctx)) {
      return err(
        returnUserError(
          new Error("Cannot update permission for SPFx project"),
          "Solution",
          SolutionError.CannotUpdatePermissionForSPFx
        )
      );
    }
    const path = `${ctx.root}/permissions.json`;
    if (!(await fs.pathExists(path))) {
      return err(
        returnSystemError(
          new Error("permissions.json is missing"),
          "Solution",
          SolutionError.MissingPermissionsJson
        )
      );
    }
    const permissionRequest = await fs.readJSON(path);
    return ok(JSON.stringify(permissionRequest));
  }

  /**
   * Checks whether solution's state is idle
   */
  private checkWhetherSolutionIsIdle(): Result<Void, FxError> {
    switch (this.runningState) {
      case SolutionRunningState.Idle:
        return ok(Void);
      case SolutionRunningState.ProvisionInProgress:
        return err(
          returnUserError(
            new Error("Provision in progress. Please wait for its completion."),
            "Solution",
            SolutionError.ProvisionInProgress
          )
        );
      case SolutionRunningState.DeployInProgress:
        return err(
          returnUserError(
            new Error("Deployment in progress. Please wait for its completion."),
            "Solution",
            SolutionError.DeploymentInProgress
          )
        );
      case SolutionRunningState.PublishInProgress:
        return err(
          returnUserError(
            new Error("Publish in progress. Please wait for its completion."),
            "Solution",
            SolutionError.PublishInProgress
          )
        );
    }
  }

  private checkWetherProvisionSucceeded(solutionConfig: SolutionConfig): boolean {
    return !!solutionConfig.get(GLOBAL_CONFIG)?.getBoolean(SOLUTION_PROVISION_SUCCEEDED);
  }

  /**
   * Provision resources. It can only run in a non-SPFx project when solution's running state is Idle.
   * Solution's provisionSucceeded config value will be set to true if provision succeeds, to false otherwise.
   *
   */
  @hooks([ErrorHandlerMW])
  async provision(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const canProvision = this.checkWhetherSolutionIsIdle();
    if (canProvision.isErr()) {
      return canProvision;
    }
    const provisioned = this.checkWetherProvisionSucceeded(ctx.config);
    if (provisioned) {
      const msg = util.format(
        getStrings().solution.AlreadyProvisionNotice,
        ctx.projectSettings?.appName
      );
      ctx.ui?.showMessage("warn", msg, false);
      const pluginCtx = getPluginContext(ctx, this.AppStudioPlugin.name);
      const remoteTeamsAppId = await this.AppStudioPlugin.provision(pluginCtx);
      if (remoteTeamsAppId.isOk()) {
        ctx.config.get(GLOBAL_CONFIG)?.set(REMOTE_TEAMS_APP_ID, remoteTeamsAppId.value);
      }
      return remoteTeamsAppId;
    }
    try {
      // Just to trigger M365 login before the concurrent execution of provision.
      // Because concurrent exectution of provision may getAccessToken() concurrently, which
      // causes 2 M365 logins before the token caching in common lib takes effect.
      await ctx.appStudioToken?.getAccessToken();

      this.runningState = SolutionRunningState.ProvisionInProgress;
      if (this.isAzureProject(ctx)) {
        const maybePermission = await this.getPermissionRequest(ctx);
        if (maybePermission.isErr()) {
          return maybePermission;
        }
        ctx.config.get(GLOBAL_CONFIG)?.set(PERMISSION_REQUEST, maybePermission.value);
      }

      const provisionResult = await this.doProvision(ctx);
      if (provisionResult.isOk()) {
        const msg = util.format(
          getStrings().solution.ProvisionSuccessNotice,
          ctx.projectSettings?.appName
        );
        ctx.logProvider?.info(msg);
        ctx.ui?.showMessage("info", msg, false);
        ctx.config.get(GLOBAL_CONFIG)?.set(SOLUTION_PROVISION_SUCCEEDED, true);
      } else {
        if (!isUserCancelError(provisionResult.error)) {
          const msg = util.format(
            getStrings().solution.ProvisionFailNotice,
            ctx.projectSettings?.appName
          );
          ctx.logProvider?.error(msg);
          ctx.config.get(GLOBAL_CONFIG)?.set(SOLUTION_PROVISION_SUCCEEDED, false);
        }
      }
      return provisionResult;
    } finally {
      this.runningState = SolutionRunningState.Idle;
      // Remove permissionRequest to prevent its persistence in config.
      ctx.config.get(GLOBAL_CONFIG)?.delete(PERMISSION_REQUEST);
    }
  }

  /**
   * provision
   */
  async doProvision(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const maybeSelectedPlugins = this.getSelectedPlugins(ctx);
    if (maybeSelectedPlugins.isErr()) {
      return maybeSelectedPlugins;
    }
    const selectedPlugins = maybeSelectedPlugins.value;

    const maybeManifest = await this.AppStudioPlugin.reloadManifestAndCheckRequiredFields(ctx.root);
    if (maybeManifest.isErr()) {
      return maybeManifest;
    }
    const manifest = maybeManifest.value;

    if (this.isAzureProject(ctx)) {
      //1. ask common questions for azure resources.
      const appName = manifest.name.short;
      const res = await fillInCommonQuestions(
        ctx,
        appName,
        ctx.config,
        ctx.dialog,
        ctx.azureAccountProvider,
        await ctx.appStudioToken?.getJsonObject()
      );
      if (res.isErr()) {
        return res;
      }
      const azureToken = await ctx.azureAccountProvider?.getAccountCredentialAsync();

      // Only Azure project requires this confirm dialog
      const username = (azureToken as any).username ? (azureToken as any).username : "";
      const subscriptionId = ctx.config.get(GLOBAL_CONFIG)?.getString("subscriptionId");
      const subscriptionName = await getSubsriptionDisplayName(azureToken!, subscriptionId!);
      const msg = util.format(
        getStrings().solution.ProvisionConfirmNotice,
        username,
        subscriptionName ? subscriptionName : subscriptionId
      );
      const confirmRes = await ctx.ui?.showMessage(
        "warn",
        msg,
        true,
        "Provision",
        "Pricing calculator"
      );
      const confirm = confirmRes?.isOk() ? confirmRes.value : undefined;

      if (confirm !== "Provision") {
        if (confirm === "Pricing calculator") {
          ctx.ui?.openUrl("https://azure.microsoft.com/en-us/pricing/calculator/");
        }
        return err(
          returnUserError(
            new Error(getStrings().solution.CancelProvision),
            "Solution",
            getStrings().solution.CancelProvision
          )
        );
      }
    }

    const pluginsWithCtx: PluginsWithContext[] = this.getPluginAndContextArray(
      ctx,
      selectedPlugins,
      manifest
    );
    const preProvisionWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.preProvision?.bind(plugin), context, plugin.name];
    });
    const provisionWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.provision?.bind(plugin), context, plugin.name];
    });
    const postProvisionWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(
      ([plugin, context]) => {
        return [plugin?.postProvision?.bind(plugin), context, plugin.name];
      }
    );

    return executeLifecycles(
      preProvisionWithCtx,
      provisionWithCtx,
      postProvisionWithCtx,
      async () => {
        ctx.logProvider?.info(
          "[Teams Toolkit]: Start provisioning. It could take several minutes."
        );
        return ok(undefined);
      },
      async (provisionResults?: any[]) => {
        ctx.logProvider?.info("[Teams Toolkit]: provison finished!");
        if (provisionWithCtx.length === provisionResults?.length) {
          provisionWithCtx.map(function (plugin, index) {
            if (plugin[2] === PluginNames.APPST) {
              ctx.config
                .get(GLOBAL_CONFIG)
                ?.set(REMOTE_TEAMS_APP_ID, provisionResults[index].value);
            }
          });
        }
        const aadPlugin = this.AadPlugin as AadAppForTeamsPlugin;
        if (selectedPlugins.some((plugin) => plugin.name === aadPlugin.name)) {
          return aadPlugin.setApplicationInContext(getPluginContext(ctx, aadPlugin.name, manifest));
        }
        return ok(undefined);
      },
      async () => {
        ctx.logProvider?.info("[Teams Toolkit]: configuration finished!");
        return ok(undefined);
      }
    );
  }

  private async canPublish(
    ctx: SolutionContext,
    manifestTpl: TeamsAppManifest
  ): Promise<Result<TeamsAppManifest, FxError>> {
    const isIdle = this.checkWhetherSolutionIsIdle();
    if (isIdle.isErr()) {
      return err(isIdle.error);
    }

    const isProvisionSucceeded = this.checkWetherProvisionSucceeded(ctx.config);
    if (!isProvisionSucceeded) {
      return err(
        returnUserError(
          new Error("Please provision before publishing"),
          "Solution",
          SolutionError.CannotPublishBeforeProvision
        )
      );
    }

    if (this.spfxSelected(ctx)) {
      const manifestString = (
        await fs.readFile(`${ctx.root}/.${ConfigFolderName}/${REMOTE_MANIFEST}`)
      ).toString();
      return ok(JSON.parse(manifestString));
    } else {
      const maybeSelectedPlugins = this.getSelectedPlugins(ctx);
      const pluginCtx = getPluginContext(ctx, this.AppStudioPlugin.name);
      return this.AppStudioPlugin.createManifestForRemote(
        pluginCtx,
        maybeSelectedPlugins,
        manifestTpl
      ).map((result) => result[1]);
    }
  }
  @hooks([ErrorHandlerMW])
  async deploy(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const isAzureProject = this.isAzureProject(ctx);
    const provisioned = this.checkWetherProvisionSucceeded(ctx.config);
    if (isAzureProject && !provisioned) {
      return err(
        returnUserError(
          new Error(
            util.format(getStrings().solution.NotProvisionedNotice, ctx.projectSettings?.appName)
          ),
          "Solution",
          SolutionError.CannotDeployBeforeProvision
        )
      );
    }
    try {
      if (this.isAzureProject(ctx)) {
        // Just to trigger M365 login before the concurrent execution of deploy.
        // Because concurrent exectution of deploy may getAccessToken() concurrently, which
        // causes 2 M365 logins before the token caching in common lib takes effect.
        await ctx.appStudioToken?.getAccessToken();
      }

      this.runningState = SolutionRunningState.DeployInProgress;
      const result = await this.doDeploy(ctx);
      if (result.isOk()) {
        if (this.isAzureProject(ctx)) {
          const msg = util.format(
            getStrings().solution.DeploySuccessNotice,
            ctx.projectSettings?.appName
          );
          ctx.logProvider?.info(msg);
          ctx.ui?.showMessage("info", msg, false);
        }
      } else {
        const msg = util.format(
          getStrings().solution.DeployFailNotice,
          ctx.projectSettings?.appName
        );
        ctx.logProvider?.info(msg);
      }

      return result;
    } finally {
      this.runningState = SolutionRunningState.Idle;
    }
  }

  /**
   * deploy
   */
  private async doDeploy(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const res = this.getSelectedPlugins(ctx);
    if (res.isErr()) {
      return res;
    }

    const loadManifestResult = await this.AppStudioPlugin.reloadManifestAndCheckRequiredFields(
      ctx.root
    );

    if (loadManifestResult.isErr()) {
      return loadManifestResult;
    }
    const manifest = loadManifestResult.value;

    const optionsToDeploy = ctx.answers![
      AzureSolutionQuestionNames.PluginSelectionDeploy
    ] as string[];
    if (optionsToDeploy === undefined || optionsToDeploy.length === 0) {
      return err(
        returnUserError(
          new Error(`No plugin selected`),
          "Solution",
          SolutionError.NoResourcePluginSelected
        )
      );
    }

    const pluginMap = getAllResourcePluginMap();
    const pluginsToDeploy: LoadedPlugin[] = [];
    for (const optionId of optionsToDeploy) {
      const filtered = pluginMap.get(optionId);
      if (filtered && res.value.find((p) => p.name === filtered.name)) {
        pluginsToDeploy.push(filtered);
      }
    }
    ctx.logProvider?.info(
      `[Solution] Selected plugins to deploy:${JSON.stringify(pluginsToDeploy.map((p) => p.name))}`
    );
    const pluginsWithCtx: PluginsWithContext[] = this.getPluginAndContextArray(
      ctx,
      pluginsToDeploy,
      manifest
    );
    const preDeployWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.preDeploy?.bind(plugin), context, plugin.name];
    });
    const deployWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.deploy?.bind(plugin), context, plugin.name];
    });
    const postDeployWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.postDeploy?.bind(plugin), context, plugin.name];
    });

    ctx.logProvider?.info(`[Solution] deploy start!`);

    return executeLifecycles(preDeployWithCtx, deployWithCtx, postDeployWithCtx);
  }
  @hooks([ErrorHandlerMW])
  async publish(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const isAzureProject = this.isAzureProject(ctx);
    const provisioned = this.checkWetherProvisionSucceeded(ctx.config);
    if (!provisioned) {
      return err(
        returnUserError(
          new Error(
            util.format(getStrings().solution.NotProvisionedNotice, ctx.projectSettings?.appName)
          ),
          "Solution",
          SolutionError.CannotPublishBeforeProvision
        )
      );
    }

    const maybeManifestTpl = await (
      this.AppStudioPlugin as AppStudioPlugin
    ).reloadManifestAndCheckRequiredFields(ctx.root);
    if (maybeManifestTpl.isErr()) {
      return err(maybeManifestTpl.error);
    }
    const manifestTpl = maybeManifestTpl.value;

    const maybeManifest = await this.canPublish(ctx, manifestTpl);
    if (maybeManifest.isErr()) {
      return maybeManifest;
    }
    const manifest = maybeManifest.value;
    try {
      this.runningState = SolutionRunningState.PublishInProgress;

      const pluginsWithCtx: PluginsWithContext[] = this.getPluginAndContextArray(
        ctx,
        [this.AppStudioPlugin],
        manifest
      );
      const publishWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
        return [plugin?.publish?.bind(plugin), context, plugin.name];
      });

      ctx.logProvider?.info(`[Solution] publish start!`);

      const results = await executeConcurrently("", publishWithCtx);

      for (const result of results) {
        if (result.isErr()) {
          const msg = util.format(
            getStrings().solution.PublishFailNotice,
            ctx.projectSettings?.appName
          );
          ctx.logProvider?.info(msg);
          return result;
        }
      }
      return ok(undefined);
    } finally {
      this.runningState = SolutionRunningState.Idle;
    }
  }

  async getTabScaffoldQuestions(
    ctx: SolutionContext,
    addAzureResource: boolean
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const tabNode = new QTreeNode({ type: "group" });

    //Frontend plugin
    const fehostPlugin: Plugin = this.FrontendPlugin;
    if (fehostPlugin.getQuestions) {
      const pluginCtx = getPluginContext(ctx, fehostPlugin.name);
      const res = await fehostPlugin.getQuestions(Stage.create, pluginCtx);
      if (res.isErr()) return res;
      if (res.value) {
        const frontendNode = res.value as QTreeNode;
        if (frontendNode.data) tabNode.addChild(frontendNode);
      }
    }

    if (addAzureResource) {
      const azureResourceNode = new QTreeNode(AzureResourcesQuestion);
      tabNode.addChild(azureResourceNode);
      const functionPlugin: Plugin = this.FunctionPlugin;
      //Azure Function
      if (functionPlugin.getQuestions) {
        const pluginCtx = getPluginContext(ctx, functionPlugin.name);
        const res = await functionPlugin.getQuestions(Stage.create, pluginCtx);
        if (res.isErr()) return res;
        if (res.value) {
          const azure_function = res.value as QTreeNode;
          azure_function.condition = { minItems: 1 };
          if (azure_function.data) azureResourceNode.addChild(azure_function);
        }
      }
      const sqlPlugin: Plugin = this.SqlPlugin;
      //Azure SQL
      if (sqlPlugin.getQuestions) {
        const pluginCtx = getPluginContext(ctx, sqlPlugin.name);
        const res = await sqlPlugin.getQuestions(Stage.create, pluginCtx);
        if (res.isErr()) return res;
        if (res.value) {
          const azure_sql = res.value as QTreeNode;
          azure_sql.condition = { contains: AzureResourceSQL.id };
          if (azure_sql.data) azureResourceNode.addChild(azure_sql);
        }
      }
    }
    return ok(tabNode);
  }

  /**
   * collect solution level question
   * @param ctx
   */
  async getQuestions(
    stage: Stage,
    ctx: SolutionContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const isDynamicQuestion = DynamicPlatforms.includes(ctx.answers!.platform!);
    const node = new QTreeNode({ type: "group" });
    let manifest: TeamsAppManifest | undefined = undefined;
    if (stage !== Stage.create && isDynamicQuestion) {
      const checkRes = this.checkWhetherSolutionIsIdle();
      if (checkRes.isErr()) return err(checkRes.error);

      const maybeManifest = await (
        this.AppStudioPlugin as AppStudioPlugin
      ).reloadManifestAndCheckRequiredFields(ctx.root);
      if (maybeManifest.isErr()) {
        return err(maybeManifest.error);
      }
      manifest = maybeManifest.value;
    }

    if (stage === Stage.create) {
      // 1. capabilities
      const capQuestion = createCapabilityQuestion();
      const capNode = new QTreeNode(capQuestion);
      node.addChild(capNode);

      // 1.1 hostType
      const hostTypeNode = new QTreeNode(FrontendHostTypeQuestion);
      hostTypeNode.condition = { contains: TabOptionItem.id };
      capNode.addChild(hostTypeNode);

      // 1.1.1 SPFX Tab
      const spfxPlugin: Plugin = new SpfxPlugin();
      if (spfxPlugin.getQuestions) {
        const pluginCtx = getPluginContext(ctx, spfxPlugin.name);
        const res = await spfxPlugin.getQuestions(Stage.create, pluginCtx);
        if (res.isErr()) return res;
        if (res.value) {
          const spfxNode = res.value as QTreeNode;
          spfxNode.condition = { equals: HostTypeOptionSPFx.id };
          if (spfxNode.data) hostTypeNode.addChild(spfxNode);
        }
      }

      // 1.1.2 Azure Tab
      const tabRes = await this.getTabScaffoldQuestions(ctx, true);
      if (tabRes.isErr()) return tabRes;
      if (tabRes.value) {
        const tabNode = tabRes.value;
        tabNode.condition = { equals: HostTypeOptionAzure.id };
        hostTypeNode.addChild(tabNode);
      }

      // 1.2 Bot
      const botPlugin: Plugin = this.BotPlugin;
      if (botPlugin.getQuestions) {
        const pluginCtx = getPluginContext(ctx, botPlugin.name);
        const res = await botPlugin.getQuestions(stage, pluginCtx);
        if (res.isErr()) return res;
        if (res.value) {
          const botGroup = res.value as QTreeNode;
          botGroup.condition = { containsAny: [BotOptionItem.id, MessageExtensionItem.id] };
          capNode.addChild(botGroup);
        }
      }

      // 1.3 Language
      const programmingLanguage = new QTreeNode(ProgrammingLanguageQuestion);
      programmingLanguage.condition = { minItems: 1 };
      capNode.addChild(programmingLanguage);
    } else if (stage === Stage.provision) {
      if (isDynamicQuestion) {
        const provisioned = this.checkWetherProvisionSucceeded(ctx.config);
        if (provisioned) return ok(undefined);
      }
      let pluginsToProvision: LoadedPlugin[];
      if (isDynamicQuestion) {
        const res = this.getSelectedPlugins(ctx);
        if (res.isErr()) {
          return err(res.error);
        }
        pluginsToProvision = res.value;
      } else {
        pluginsToProvision = getAllResourcePlugins();
      }
      if (!isDynamicQuestion) {
        node.addChild(new QTreeNode(AskSubscriptionQuestion));
      }
      for (const plugin of pluginsToProvision) {
        if (plugin.getQuestions) {
          const pluginCtx = getPluginContext(ctx, plugin.name, manifest);
          const getQuestionRes = await plugin.getQuestions(stage, pluginCtx);
          if (getQuestionRes.isErr()) return getQuestionRes;
          if (getQuestionRes.value) {
            const subnode = getQuestionRes.value as QTreeNode;
            node.addChild(subnode);
          }
        }
      }
    } else if (stage === Stage.deploy) {
      if (isDynamicQuestion) {
        const isAzureProject = this.isAzureProject(ctx);
        const provisioned = this.checkWetherProvisionSucceeded(ctx.config);
        if (isAzureProject && !provisioned) {
          return err(
            returnUserError(
              new Error(getStrings().solution.FailedToDeployBeforeProvision),
              "Solution",
              SolutionError.CannotDeployBeforeProvision
            )
          );
        }
      }
      let pluginsToDeploy: LoadedPlugin[];
      if (isDynamicQuestion) {
        const res = this.getSelectedPlugins(ctx);
        if (res.isErr()) {
          return err(
            returnUserError(
              new Error("No resource to deploy"),
              "Solution",
              SolutionError.NoResourceToDeploy
            )
          );
        }
        pluginsToDeploy = res.value.filter((plugin) => !!plugin.deploy);
      } else {
        const allPlugins = getAllResourcePlugins();
        pluginsToDeploy = allPlugins.filter((plugin) => !!plugin.deploy);
      }

      if (pluginsToDeploy.length === 0) {
        return err(
          returnUserError(
            new Error("No resource to deploy"),
            "Solution",
            SolutionError.NoResourceToDeploy
          )
        );
      }
      const pluginPrefix = "fx-resource-";
      const options: OptionItem[] = pluginsToDeploy.map((plugin) => {
        const item: OptionItem = {
          id: plugin.name,
          label: plugin.displayName,
          cliName: plugin.name.replace(pluginPrefix, ""),
        };
        return item;
      });

      const selectQuestion = DeployPluginSelectQuestion;
      selectQuestion.staticOptions = options;
      selectQuestion.default = options.map((i) => i.id);
      const pluginSelection = new QTreeNode(selectQuestion);
      node.addChild(pluginSelection);

      for (const plugin of pluginsToDeploy) {
        if (plugin.getQuestions) {
          const pluginCtx = getPluginContext(ctx, plugin.name, manifest);
          const getQuestionRes = await plugin.getQuestions(stage, pluginCtx);
          if (getQuestionRes.isErr()) return getQuestionRes;
          if (getQuestionRes.value) {
            const subnode = getQuestionRes.value as QTreeNode;
            subnode.condition = { contains: plugin.name };
            if (subnode.data) pluginSelection.addChild(subnode);
          }
        }
      }
    } else if (stage === Stage.publish) {
      if (isDynamicQuestion) {
        const isAzureProject = this.isAzureProject(ctx);
        const provisioned = this.checkWetherProvisionSucceeded(ctx.config);
        if (isAzureProject && !provisioned) {
          return err(
            returnUserError(
              new Error(getStrings().solution.FailedToPublishBeforeProvision),
              "Solution",
              SolutionError.CannotPublishBeforeProvision
            )
          );
        }
        if (!provisioned && this.spfxSelected(ctx)) {
          if (ctx.answers?.platform === Platform.VSCode) {
            ctx.ui?.showMessage(
              "error",
              getStrings().solution.SPFxAskProvisionBeforePublish,
              false
            );
            throw CancelError;
          } else {
            return err(
              returnUserError(
                new Error(getStrings().solution.SPFxAskProvisionBeforePublish),
                "Solution",
                SolutionError.CannotPublishBeforeProvision
              )
            );
          }
        }
      }
      const pluginsToPublish = [this.AppStudioPlugin];
      for (const plugin of pluginsToPublish) {
        const pluginCtx = getPluginContext(ctx, plugin.name, manifest);
        if (plugin.getQuestions) {
          const getQuestionRes = await plugin.getQuestions(stage, pluginCtx);
          if (getQuestionRes.isErr()) return getQuestionRes;
          if (getQuestionRes.value) {
            const subnode = getQuestionRes.value as QTreeNode;
            node.addChild(subnode);
          }
        }
      }
    }
    return ok(node);
  }

  async localDebug(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const maybePermission = await this.getPermissionRequest(ctx);
    if (maybePermission.isErr()) {
      return maybePermission;
    }
    try {
      ctx.config.get(GLOBAL_CONFIG)?.set(PERMISSION_REQUEST, maybePermission.value);
      const result = await this.doLocalDebug(ctx);
      return result;
    } finally {
      ctx.config.get(GLOBAL_CONFIG)?.delete(PERMISSION_REQUEST);
    }
  }

  async doLocalDebug(ctx: SolutionContext): Promise<Result<any, FxError>> {
    const maybeSelectedPlugins = this.getSelectedPlugins(ctx);

    if (maybeSelectedPlugins.isErr()) {
      return maybeSelectedPlugins;
    }

    const selectedPlugins = maybeSelectedPlugins.value;

    const maybeManifest = await (
      this.AppStudioPlugin as AppStudioPlugin
    ).reloadManifestAndCheckRequiredFields(ctx.root);
    if (maybeManifest.isErr()) {
      return err(maybeManifest.error);
    }
    const manifest = maybeManifest.value;

    // Just to trigger M365 login before the concurrent execution of localDebug.
    // Because concurrent exectution of localDebug may getAccessToken() concurrently, which
    // causes 2 M365 logins before the token caching in common lib takes effect.
    await ctx.appStudioToken?.getAccessToken();

    const pluginsWithCtx: PluginsWithContext[] = this.getPluginAndContextArray(
      ctx,
      selectedPlugins,
      manifest
    );
    const localDebugWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(([plugin, context]) => {
      return [plugin?.localDebug?.bind(plugin), context, plugin.name];
    });
    const postLocalDebugWithCtx: LifecyclesWithContext[] = pluginsWithCtx.map(
      ([plugin, context]) => {
        return [plugin?.postLocalDebug?.bind(plugin), context, plugin.name];
      }
    );

    const localDebugResults = await executeConcurrently("", localDebugWithCtx);
    for (const localDebugResult of localDebugResults) {
      if (localDebugResult.isErr()) {
        return localDebugResult;
      }
    }

    const aadPlugin = this.AadPlugin as AadAppForTeamsPlugin;
    if (selectedPlugins.some((plugin) => plugin.name === aadPlugin.name)) {
      const result = aadPlugin.setApplicationInContext(
        getPluginContext(ctx, aadPlugin.name, manifest),
        true
      );
      if (result.isErr()) {
        return result;
      }
    }
    const result = this.loadTeamsAppTenantId(ctx.config, await ctx.appStudioToken?.getJsonObject());

    if (result.isErr()) {
      return result;
    }

    const postLocalDebugResults = await executeConcurrently("post", postLocalDebugWithCtx);
    for (const postLocalDebugResult of postLocalDebugResults) {
      if (postLocalDebugResult.isErr()) {
        return postLocalDebugResult;
      }
    }

    const localTeamsAppID = ctx.config.get(GLOBAL_CONFIG)?.getString(LOCAL_DEBUG_TEAMS_APP_ID);

    const appStudioPlugin = this.AppStudioPlugin as AppStudioPlugin;
    const maybeTeamsAppId = await appStudioPlugin.getAppDefinitionAndUpdate(
      getPluginContext(ctx, this.AppStudioPlugin.name, manifest),
      "localDebug",
      manifest
    );
    if (maybeTeamsAppId.isErr()) {
      return maybeTeamsAppId;
    }
    if (!localTeamsAppID) {
      ctx.config.get(GLOBAL_CONFIG)?.set(LOCAL_DEBUG_TEAMS_APP_ID, maybeTeamsAppId.value);
    }

    return ok(Void);
  }

  private parseTeamsAppTenantId(appStudioToken?: object): Result<string, FxError> {
    if (appStudioToken === undefined) {
      return err(
        returnSystemError(
          new Error("Graph token json is undefined"),
          "Solution",
          SolutionError.NoAppStudioToken
        )
      );
    }

    const teamsAppTenantId = (appStudioToken as any).tid;
    if (
      teamsAppTenantId === undefined ||
      !(typeof teamsAppTenantId === "string") ||
      teamsAppTenantId.length === 0
    ) {
      return err(
        returnSystemError(
          new Error("Cannot find teams app tenant id"),
          "Solution",
          SolutionError.NoTeamsAppTenantId
        )
      );
    }
    return ok(teamsAppTenantId);
  }

  private loadTeamsAppTenantId(
    config: SolutionConfig,
    appStudioToken?: object
  ): Result<SolutionConfig, FxError> {
    return this.parseTeamsAppTenantId(appStudioToken).andThen((teamsAppTenantId) => {
      config.get(GLOBAL_CONFIG)?.set("teamsAppTenantId", teamsAppTenantId);
      return ok(config);
    });
  }

  getAzureSolutionSettings(ctx: SolutionContext): AzureSolutionSettings {
    return ctx.projectSettings?.solutionSettings as AzureSolutionSettings;
  }

  async getQuestionsForAddResource(
    func: Func,
    ctx: SolutionContext,
    manifest?: TeamsAppManifest
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const isDynamicQuestion = DynamicPlatforms.includes(ctx.answers!.platform!);
    const settings = this.getAzureSolutionSettings(ctx);

    if (
      isDynamicQuestion &&
      !(
        settings.hostType === HostTypeOptionAzure.id &&
        settings.capabilities &&
        settings.capabilities.includes(TabOptionItem.id)
      )
    ) {
      return err(
        returnUserError(
          new Error("Add resource is only supported for Tab app hosted in Azure."),
          "Solution",
          SolutionError.AddResourceNotSupport
        )
      );
    }

    const selectedPlugins = settings.activeResourcePlugins || [];

    if (!selectedPlugins) {
      return err(
        returnUserError(
          new Error("selectedPlugins is empty"),
          "Solution",
          SolutionError.InternelError
        )
      );
    }
    const functionPlugin: Plugin = this.FunctionPlugin;
    const sqlPlugin: Plugin = this.SqlPlugin;
    const apimPlugin: Plugin = this.ApimPlugin;
    const alreadyHaveFunction = selectedPlugins.includes(functionPlugin.name);
    const alreadyHaveSQL = selectedPlugins.includes(sqlPlugin.name);
    const alreadyHaveAPIM = selectedPlugins.includes(apimPlugin.name);

    const addQuestion = createAddAzureResourceQuestion(
      alreadyHaveFunction,
      alreadyHaveSQL,
      alreadyHaveAPIM
    );

    const addAzureResourceNode = new QTreeNode(addQuestion);

    // there two cases to add function re-scaffold: 1. select add function   2. select add sql and function is not selected when creating
    if (functionPlugin.getQuestionsForUserTask) {
      const pluginCtx = getPluginContext(ctx, functionPlugin.name, manifest);
      const res = await functionPlugin.getQuestionsForUserTask(func, pluginCtx);
      if (res.isErr()) return res;
      if (res.value) {
        const azure_function = res.value as QTreeNode;
        if (alreadyHaveFunction) {
          // if already has function, the question will appear depends on whether user select function, otherwise, the question will always show
          azure_function.condition = { contains: AzureResourceFunction.id };
        } else {
          // if not function activated, select any option will trigger function question
          azure_function.condition = { minItems: 1 };
        }
        if (azure_function.data) addAzureResourceNode.addChild(azure_function);
      }
    }

    //Azure SQL
    if (sqlPlugin.getQuestionsForUserTask && !alreadyHaveSQL) {
      const pluginCtx = getPluginContext(ctx, sqlPlugin.name, manifest);
      const res = await sqlPlugin.getQuestionsForUserTask(func, pluginCtx);
      if (res.isErr()) return res;
      if (res.value) {
        const azure_sql = res.value as QTreeNode;
        azure_sql.condition = { contains: AzureResourceSQL.id };
        if (azure_sql.data) addAzureResourceNode.addChild(azure_sql);
      }
    }

    //APIM
    if (apimPlugin.getQuestionsForUserTask && (!alreadyHaveAPIM || !isDynamicQuestion)) {
      const pluginCtx = getPluginContext(ctx, apimPlugin.name, manifest);
      const res = await apimPlugin.getQuestionsForUserTask(func, pluginCtx);
      if (res.isErr()) return res;
      if (res.value) {
        const groupNode = new QTreeNode({ type: "group" });
        groupNode.condition = { contains: AzureResourceApim.id };
        addAzureResourceNode.addChild(groupNode);
        const apim = res.value as QTreeNode;
        if (apim.data) {
          const funcNode = new QTreeNode(AskSubscriptionQuestion);
          AskSubscriptionQuestion.func = async (
            inputs: Inputs
          ): Promise<Result<SubscriptionInfo, FxError>> => {
            const res = await checkSubscription(ctx);
            if (res.isOk()) {
              const sub = res.value;
              inputs.subscriptionId = sub.subscriptionId;
              inputs.tenantId = sub.tenantId;
            }
            return res;
          };
          groupNode.addChild(funcNode);
          groupNode.addChild(apim);
        }
      }
    }
    return ok(addAzureResourceNode);
  }

  async getQuestionsForAddCapability(
    ctx: SolutionContext,
    manifest?: TeamsAppManifest
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const isDynamicQuestion = DynamicPlatforms.includes(ctx.answers!.platform!);
    const settings = this.getAzureSolutionSettings(ctx);

    if (!(settings.hostType === HostTypeOptionAzure.id) && isDynamicQuestion) {
      return err(
        returnUserError(
          new Error("Add capability is not supported for SPFx project"),
          "Solution",
          SolutionError.AddResourceNotSupport
        )
      );
    }

    const capabilities = settings.capabilities || [];

    const alreadyHaveTab = capabilities.includes(TabOptionItem.id);

    const alreadyHaveBotOrMe =
      capabilities.includes(BotOptionItem.id) || capabilities.includes(MessageExtensionItem.id);

    if (alreadyHaveBotOrMe && alreadyHaveTab) {
      const cannotAddCapWarnMsg =
        "Your App already has both Tab and Bot/Me, can not Add Capability.";
      ctx.ui?.showMessage("error", cannotAddCapWarnMsg, false);
      return ok(undefined);
    }

    const addCapQuestion = addCapabilityQuestion(alreadyHaveTab, alreadyHaveBotOrMe);

    const addCapNode = new QTreeNode(addCapQuestion);

    //Tab sub tree
    if (!alreadyHaveTab || !isDynamicQuestion) {
      const tabRes = await this.getTabScaffoldQuestions(ctx, false);
      if (tabRes.isErr()) return tabRes;
      if (tabRes.value) {
        const tabNode = tabRes.value;
        tabNode.condition = { contains: TabOptionItem.id };
        addCapNode.addChild(tabNode);
      }
    }

    //Bot sub tree
    const botPlugin: Plugin = this.BotPlugin;
    if ((!alreadyHaveBotOrMe || !isDynamicQuestion) && botPlugin.getQuestions) {
      const pluginCtx = getPluginContext(ctx, botPlugin.name, manifest);
      const res = await botPlugin.getQuestions(Stage.create, pluginCtx);
      if (res.isErr()) return res;
      if (res.value) {
        const child = res.value as QTreeNode;
        child.condition = { containsAny: [BotOptionItem.id, MessageExtensionItem.id] };
        if (child.data) addCapNode.addChild(child);
      }
    }

    return ok(addCapNode);
  }

  /**
   * user questions for customized task
   */
  async getQuestionsForUserTask(
    func: Func,
    ctx: SolutionContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    const isDynamicQuestion = DynamicPlatforms.includes(ctx.answers!.platform!);
    const namespace = func.namespace;
    const array = namespace.split("/");
    let manifest: TeamsAppManifest | undefined = undefined;
    if (isDynamicQuestion) {
      const maybeManifest = await (
        this.AppStudioPlugin as AppStudioPlugin
      ).reloadManifestAndCheckRequiredFields(ctx.root);
      if (maybeManifest.isErr()) {
        return err(maybeManifest.error);
      }
      manifest = maybeManifest.value;
    }
    if (func.method === "addCapability") {
      return await this.getQuestionsForAddCapability(ctx, manifest);
    }
    if (func.method === "addResource") {
      return await this.getQuestionsForAddResource(func, ctx, manifest);
    }
    if (array.length == 2) {
      const pluginName = array[1];
      const pluginMap = getAllResourcePluginMap();
      const plugin = pluginMap.get(pluginName);
      if (plugin) {
        if (plugin.getQuestionsForUserTask) {
          const pctx = getPluginContext(ctx, plugin.name, manifest);
          return await plugin.getQuestionsForUserTask(func, pctx);
        } else {
          return ok(undefined);
        }
      }
    }
    return ok(undefined);
  }

  async executeAddResource(ctx: SolutionContext): Promise<Result<any, FxError>> {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResourceStart, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
    });

    if (!ctx.answers) {
      return err(
        returnUserError(new Error(`answer is empty!`), "Solution", SolutionError.InternelError)
      );
    }
    const settings = this.getAzureSolutionSettings(ctx);
    const originalSettings = deepCopy(settings);
    if (
      !(
        settings.hostType === HostTypeOptionAzure.id &&
        settings.capabilities &&
        settings.capabilities.includes(TabOptionItem.id)
      )
    ) {
      const e = returnUserError(
        new Error("Add resource is only supported for Tab app hosted in Azure."),
        "Solution",
        SolutionError.AddResourceNotSupport
      );

      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddResource,
          e,
          ctx.telemetryReporter
        )
      );
    }
    const selectedPlugins = settings.activeResourcePlugins;
    const functionPlugin: Plugin = this.FunctionPlugin;
    const sqlPlugin: Plugin = this.SqlPlugin;
    const apimPlugin: Plugin = this.ApimPlugin;
    const alreadyHaveFunction = selectedPlugins?.includes(functionPlugin.name);
    const alreadyHaveSql = selectedPlugins?.includes(sqlPlugin.name);
    const alreadyHaveApim = selectedPlugins?.includes(apimPlugin.name);

    const addResourcesAnswer = ctx.answers[AzureSolutionQuestionNames.AddResources] as string[];

    if (!addResourcesAnswer) {
      return err(
        returnUserError(
          new Error(`answer of ${AzureSolutionQuestionNames.AddResources} is empty!`),
          "Solution",
          SolutionError.InvalidInput
        )
      );
    }

    const addSQL = addResourcesAnswer.includes(AzureResourceSQL.id);
    const addFunc = addResourcesAnswer.includes(AzureResourceFunction.id);
    const addApim = addResourcesAnswer.includes(AzureResourceApim.id);

    if ((alreadyHaveSql && addSQL) || (alreadyHaveApim && addApim)) {
      const e = returnUserError(
        new Error("SQL/APIM is already added."),
        "Solution",
        SolutionError.AddResourceNotSupport
      );
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddResource,
          e,
          ctx.telemetryReporter
        )
      );
    }

    let addNewResoruceToProvision = false;
    const notifications: string[] = [];
    const pluginsToScaffold: LoadedPlugin[] = [this.LocalDebugPlugin];
    const azureResource = settings.azureResources || [];
    if (addFunc || ((addSQL || addApim) && !alreadyHaveFunction)) {
      pluginsToScaffold.push(functionPlugin);
      if (!azureResource.includes(AzureResourceFunction.id)) {
        azureResource.push(AzureResourceFunction.id);
        addNewResoruceToProvision = true;
      }
      notifications.push(AzureResourceFunction.label);
    }
    if (addSQL && !alreadyHaveSql) {
      pluginsToScaffold.push(sqlPlugin);
      azureResource.push(AzureResourceSQL.id);
      notifications.push(AzureResourceSQL.label);
      addNewResoruceToProvision = true;
    }
    if (addApim && !alreadyHaveApim) {
      pluginsToScaffold.push(apimPlugin);
      azureResource.push(AzureResourceApim.id);
      notifications.push(AzureResourceApim.label);
      addNewResoruceToProvision = true;
    }

    if (notifications.length > 0) {
      await this.reloadPlugins(settings);
      ctx.logProvider?.info(`start scaffolding ${notifications.join(",")}.....`);
      const scaffoldRes = await this.doScaffold(ctx, pluginsToScaffold);
      if (scaffoldRes.isErr()) {
        ctx.logProvider?.info(`failed to scaffold ${notifications.join(",")}!`);
        ctx.projectSettings!.solutionSettings = originalSettings;
        return err(
          sendErrorTelemetryThenReturnError(
            SolutionTelemetryEvent.AddResource,
            scaffoldRes.error,
            ctx.telemetryReporter
          )
        );
      }
      ctx.logProvider?.info(`finish scaffolding ${notifications.join(",")}!`);
      if (addNewResoruceToProvision)
        ctx.config.get(GLOBAL_CONFIG)?.set(SOLUTION_PROVISION_SUCCEEDED, false); //if selected plugin changed, we need to re-do provision
      ctx.ui?.showMessage(
        "info",
        util.format(
          ctx.answers.platform === Platform.CLI
            ? getStrings().solution.AddResourceNoticeForCli
            : getStrings().solution.AddResourceNotice,
          notifications.join(",")
        ),
        false
      );
    }

    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddResource, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Resources]: addResourcesAnswer.join(";"),
    });
    return ok(Void);
  }

  async executeAddCapability(func: Func, ctx: SolutionContext): Promise<Result<any, FxError>> {
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapabilityStart, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
    });
    if (!ctx.answers) {
      return err(
        returnUserError(new Error(`answer is empty!`), "Solution", SolutionError.InternelError)
      );
    }
    const settings = this.getAzureSolutionSettings(ctx);
    const originalSettings = deepCopy(settings);
    if (!(settings.hostType === HostTypeOptionAzure.id)) {
      const e = returnUserError(
        new Error("Add capability is not supported for SPFx project"),
        "Solution",
        SolutionError.FailedToAddCapability
      );
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddCapability,
          e,
          ctx.telemetryReporter
        )
      );
    }

    const capabilitiesAnswer = ctx.answers[AzureSolutionQuestionNames.Capabilities] as string[];
    if (!capabilitiesAnswer || capabilitiesAnswer.length === 0) {
      ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
        [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
        [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
        [SolutionTelemetryProperty.Capabilities]: [].join(";"),
      });
      return ok(Void);
    }

    if (
      (settings.capabilities?.includes(BotOptionItem.id) ||
        settings.capabilities?.includes(MessageExtensionItem.id)) &&
      (capabilitiesAnswer.includes(BotOptionItem.id) ||
        capabilitiesAnswer.includes(MessageExtensionItem.id))
    ) {
      const e = returnUserError(
        new Error("Application already contains a Bot and/or Messaging Extension"),
        "Solution",
        SolutionError.FailedToAddCapability
      );
      return err(
        sendErrorTelemetryThenReturnError(
          SolutionTelemetryEvent.AddCapability,
          e,
          ctx.telemetryReporter
        )
      );
    }
    let change = false;
    const notifications: string[] = [];
    const pluginsToScaffold: LoadedPlugin[] = [this.LocalDebugPlugin];
    for (const cap of capabilitiesAnswer!) {
      if (!settings.capabilities.includes(cap)) {
        settings.capabilities.push(cap);
        change = true;
        if (cap === TabOptionItem.id) {
          notifications.push("Azure Tab Frontend");
          pluginsToScaffold.push(this.FrontendPlugin);
        } else if (
          (cap === BotOptionItem.id || cap === MessageExtensionItem.id) &&
          !pluginsToScaffold.includes(this.BotPlugin)
        ) {
          notifications.push("Bot/MessageExtension");
          pluginsToScaffold.push(this.BotPlugin);
        }
      }
    }

    if (change) {
      await this.reloadPlugins(settings);
      if (this.isAzureProject(ctx)) {
        const manifest = await (this.AppStudioPlugin as AppStudioPlugin).createManifest(
          ctx.projectSettings!
        );
        // if (manifest) Object.assign(ctx.app, manifest);
        await fs.writeFile(
          `${ctx.root}/.${ConfigFolderName}/${REMOTE_MANIFEST}`,
          JSON.stringify(manifest, null, 4)
        );
        await fs.writeJSON(`${ctx.root}/permissions.json`, DEFAULT_PERMISSION_REQUEST, {
          spaces: 4,
        });
      } else {
        const manifest = await new SpfxPlugin().getManifest();
        await fs.writeFile(
          `${ctx.root}/.${ConfigFolderName}/${REMOTE_MANIFEST}`,
          JSON.stringify(manifest, null, 4)
        );
      }
      ctx.logProvider?.info(`start scaffolding ${notifications.join(",")}.....`);
      const scaffoldRes = await this.doScaffold(ctx, pluginsToScaffold);
      if (scaffoldRes.isErr()) {
        ctx.logProvider?.info(`failed to scaffold ${notifications.join(",")}!`);
        ctx.projectSettings!.solutionSettings = originalSettings;
        return err(
          sendErrorTelemetryThenReturnError(
            SolutionTelemetryEvent.AddCapability,
            scaffoldRes.error,
            ctx.telemetryReporter
          )
        );
      }
      ctx.logProvider?.info(`finish scaffolding ${notifications.join(",")}!`);
      ctx.config.get(GLOBAL_CONFIG)?.set(SOLUTION_PROVISION_SUCCEEDED, false);
      const msg = util.format(
        ctx.answers.platform === Platform.CLI
          ? getStrings().solution.AddCapabilityNoticeForCli
          : getStrings().solution.AddCapabilityNotice,
        notifications.join(",")
      );
      ctx.ui?.showMessage("info", msg, false);

      ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
        [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
        [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
        [SolutionTelemetryProperty.Capabilities]: capabilitiesAnswer.join(";"),
      });
      return ok({});
    }
    const cannotAddCapWarnMsg = "Add nothing";
    ctx.ui?.showMessage("warn", cannotAddCapWarnMsg, false);
    ctx.telemetryReporter?.sendTelemetryEvent(SolutionTelemetryEvent.AddCapability, {
      [SolutionTelemetryProperty.Component]: SolutionTelemetryComponentName,
      [SolutionTelemetryProperty.Success]: SolutionTelemetrySuccess.Yes,
      [SolutionTelemetryProperty.Capabilities]: [].join(";"),
    });
    return ok({});
  }
  /**
   * execute user task
   */
  @hooks([ErrorHandlerMW])
  async executeUserTask(func: Func, ctx: SolutionContext): Promise<Result<any, FxError>> {
    if (!ctx.answers)
      return err(
        returnUserError(new Error(`answer is empty!`), "Solution", SolutionError.InternelError)
      );
    const namespace = func.namespace;
    const method = func.method;
    const array = namespace.split("/");
    if (method === "addCapability") {
      return this.executeAddCapability(func, ctx!);
    }
    if (method === "addResource") {
      return this.executeAddResource(ctx);
    }
    if (namespace.includes("solution")) {
      if (method === "registerTeamsAppAndAad") {
        const maybeParams = this.extractParamForRegisterTeamsAppAndAad(ctx.answers);
        if (maybeParams.isErr()) {
          return maybeParams;
        }
        return this.registerTeamsAppAndAad(ctx, maybeParams.value);
      } else if (method === "VSpublish") {
        // VSpublish means VS calling cli to do publish. It is different than normal cli work flow
        // It's teamsfx init followed by teamsfx  publish without running provision.
        // Using executeUserTask here could bypass the fx project check.
        if (ctx.answers?.platform !== "vs") {
          return err(
            returnSystemError(
              new Error(`VS publish is not supposed to run on platform ${ctx.answers?.platform}`),
              "Solution",
              SolutionError.UnsupportedPlatform
            )
          );
        }
        const appStudioPlugin = this.AppStudioPlugin as AppStudioPlugin;
        const pluginCtx = getPluginContext(ctx, appStudioPlugin.name);
        return appStudioPlugin.publish(pluginCtx);
      } else if (method === "validateManifest") {
        const appStudioPlugin = this.AppStudioPlugin as AppStudioPlugin;
        const pluginCtx = getPluginContext(ctx, appStudioPlugin.name);
        return await appStudioPlugin.validateManifest(pluginCtx);
      } else if (method === "buildPackage") {
        const appStudioPlugin = this.AppStudioPlugin as AppStudioPlugin;
        const pluginCtx = getPluginContext(ctx, appStudioPlugin.name);
        return await appStudioPlugin.buildTeamsPackage(
          pluginCtx,
          `${ctx.root}/.${ConfigFolderName}`
        );
      } else if (array.length == 2) {
        const pluginName = array[1];
        const pluginMap = getAllResourcePluginMap();
        const plugin = pluginMap.get(pluginName);
        if (plugin && plugin.executeUserTask) {
          const maybeManifest = await (
            this.AppStudioPlugin as AppStudioPlugin
          ).reloadManifestAndCheckRequiredFields(ctx.root);
          if (maybeManifest.isErr()) {
            return maybeManifest;
          }
          const manifestTpl = maybeManifest.value;
          const pctx = getPluginContext(ctx, plugin.name, manifestTpl);
          return plugin.executeUserTask(func, pctx);
        }
      }
    }

    return err(
      returnUserError(
        new Error(`executeUserTaskRouteFailed:${JSON.stringify(func)}`),
        "Solution",
        `executeUserTaskRouteFailed`
      )
    );
  }

  private extractParamForRegisterTeamsAppAndAad(
    answers?: Inputs
  ): Result<ParamForRegisterTeamsAppAndAad, FxError> {
    if (answers == undefined) {
      return err(
        returnSystemError(
          new Error("Input is undefined"),
          "Solution",
          SolutionError.FailedToGetParamForRegisterTeamsAppAndAad
        )
      );
    }

    const param: ParamForRegisterTeamsAppAndAad = {
      "app-name": "",
      endpoint: "",
      environment: "local",
      "root-path": "",
    };
    for (const key of Object.keys(param)) {
      const value = answers[key];
      if (value == undefined) {
        return err(
          returnSystemError(
            new Error(`${key} not found`),
            "Solution",
            SolutionError.FailedToGetParamForRegisterTeamsAppAndAad
          )
        );
      }
      (param as any)[key] = value;
    }

    return ok(param);
  }

  private prepareConfigForRegisterTeamsAppAndAad(
    config: SolutionConfig,
    params: ParamForRegisterTeamsAppAndAad
  ): string {
    const endpoint = params.endpoint;
    const domain = new URL(endpoint).hostname;

    if (config.get(GLOBAL_CONFIG) == undefined) {
      config.set(GLOBAL_CONFIG, new ConfigMap());
    }
    config.get(GLOBAL_CONFIG)!.set(PERMISSION_REQUEST, JSON.stringify(DEFAULT_PERMISSION_REQUEST));
    const aadPlugin = this.AadPlugin;
    if (config.get(aadPlugin.name) == undefined) {
      config.set(aadPlugin.name, new ConfigMap());
    }
    config.get(aadPlugin.name)!.set("domain", domain);
    config.get(aadPlugin.name)!.set("endpoint", endpoint);
    return domain;
  }

  private extractConfigForRegisterTeamsAppAndAad(
    config: SolutionConfig,
    isLocal: boolean
  ): Result<{ aadId: string; applicationIdUri: string; clientSecret: string }, FxError> {
    const aadPlugin = this.AadPlugin;
    const aadId = config.get(aadPlugin.name)?.get(isLocal ? LOCAL_DEBUG_AAD_ID : REMOTE_AAD_ID);
    if (aadId === undefined || typeof aadId !== "string") {
      return err(
        returnSystemError(
          new Error(`config ${LOCAL_DEBUG_AAD_ID} is missing`),
          "Solution",
          SolutionError.RegisterTeamsAppAndAadError
        )
      );
    }
    const applicationIdUri = config
      .get(aadPlugin.name)
      ?.get(isLocal ? LOCAL_APPLICATION_ID_URIS : REMOTE_APPLICATION_ID_URIS);
    if (applicationIdUri === undefined || typeof applicationIdUri !== "string") {
      return err(
        returnSystemError(
          new Error(`config ${LOCAL_APPLICATION_ID_URIS} is missing`),
          "Solution",
          SolutionError.RegisterTeamsAppAndAadError
        )
      );
    }
    const clientSecret = config
      .get(aadPlugin.name)
      ?.get(isLocal ? LOCAL_CLIENT_SECRET : REMOTE_CLIENT_SECRET);
    if (clientSecret === undefined || typeof clientSecret !== "string") {
      return err(
        returnSystemError(
          new Error(`config ${LOCAL_CLIENT_SECRET} is missing`),
          "Solution",
          SolutionError.RegisterTeamsAppAndAadError
        )
      );
    }
    return ok({
      aadId,
      applicationIdUri,
      clientSecret,
    });
  }

  /**
   * This function is only called by cli: teamsfx init. The context may be different from that of vsc: no .${ConfigFolderName} folder, no permissions.json
   * In order to reuse aad plugin, we need to pretend we are still in vsc context. Currently, we don't support icons, because icons are not included in the
   * current contract.
   */
  private async registerTeamsAppAndAad(
    ctx: SolutionContext,
    params: ParamForRegisterTeamsAppAndAad
  ): Promise<
    Result<
      {
        teamsAppId: string;
        clientId: string;
        clientSecret: string;
        tenantId: string;
        applicationIdUri: string;
      },
      FxError
    >
  > {
    const rootPath = params["root-path"];
    const isLocal: boolean = params.environment === "local";
    const mockedManifest = new TeamsAppManifest();
    mockedManifest.name.short = params["app-name"];
    const domain = this.prepareConfigForRegisterTeamsAppAndAad(ctx.config, params);
    const aadPlugin = this.AadPlugin as AadAppForTeamsPlugin;
    const aadPluginCtx = getPluginContext(ctx, aadPlugin.name, mockedManifest);

    const provisionResult = isLocal
      ? await aadPlugin.localDebug(aadPluginCtx)
      : await aadPlugin.provision(aadPluginCtx);
    if (provisionResult.isErr()) {
      return provisionResult;
    }
    aadPlugin.setApplicationInContext(aadPluginCtx, isLocal);
    const postProvisionResult = isLocal
      ? await aadPlugin.postLocalDebug(aadPluginCtx)
      : await aadPlugin.postProvision(aadPluginCtx);
    if (postProvisionResult.isErr()) {
      return postProvisionResult;
    }

    const configResult = this.extractConfigForRegisterTeamsAppAndAad(ctx.config, isLocal);
    if (configResult.isErr()) {
      return err(configResult.error);
    }

    const manifestPath: string = path.join(
      rootPath,
      "manifest",
      isLocal ? "local" : "remote",
      "manifest.json"
    );
    const appSettingsJSONPath = path.join(
      rootPath,
      isLocal ? "appsettings.Development.json" : "appsettings.json"
    );

    const manifestTpl = (await fs.readFile(manifestPath)).toString();
    const manifestStr: string = Mustache.render(manifestTpl, {
      "client-id": configResult.value.aadId,
      "app-name": params["app-name"],
      endpoint: params.endpoint,
      domain: domain,
      "application-id-uri": configResult.value.applicationIdUri,
    });
    const manifest: TeamsAppManifest = JSON.parse(manifestStr);
    await fs.writeFile(manifestPath, manifestStr);
    const appStudioPlugin: AppStudioPlugin = this.AppStudioPlugin as any;
    const maybeTeamsAppId = await appStudioPlugin.getAppDefinitionAndUpdate(
      getPluginContext(ctx, this.AppStudioPlugin.name, manifest),
      "remote",
      manifest
    );
    if (maybeTeamsAppId.isErr()) {
      return err(maybeTeamsAppId.error);
    }
    const teamsAppId = maybeTeamsAppId.value;

    const appSettingsJSONTpl = (await fs.readFile(appSettingsJSONPath)).toString();
    const maybeTenantId = this.parseTeamsAppTenantId(await ctx.appStudioToken?.getJsonObject());
    if (maybeTenantId.isErr()) {
      return err(maybeTenantId.error);
    }
    const appSettingsJSON = Mustache.render(appSettingsJSONTpl, {
      "client-id": configResult.value.aadId,
      "client-secret": configResult.value.clientSecret,
      "application-id-uri": configResult.value.applicationIdUri,
      endpoint: params.endpoint,
      "tenant-id": maybeTenantId.value,
    });
    await fs.writeFile(appSettingsJSONPath, appSettingsJSON);

    if (isLocal) {
      const launchSettingsJSONPath: string = path.join(
        rootPath,
        "Properties",
        "launchSettings.json"
      );
      const launchSettingsJSONTpl = (await fs.readFile(launchSettingsJSONPath)).toString();
      const launchSettingsJSON = Mustache.render(launchSettingsJSONTpl, {
        "teams-app-id": teamsAppId,
      });
      await fs.writeFile(launchSettingsJSONPath, launchSettingsJSON);
    }
    // Remove permissionRequest to prevent its persistence in config.
    ctx.config.get(GLOBAL_CONFIG)?.delete(PERMISSION_REQUEST);
    return ok({
      teamsAppId: teamsAppId,
      clientId: configResult.value.aadId,
      clientSecret: configResult.value.clientSecret,
      tenantId: maybeTenantId.value,
      applicationIdUri: configResult.value.applicationIdUri,
    });
  }
}
