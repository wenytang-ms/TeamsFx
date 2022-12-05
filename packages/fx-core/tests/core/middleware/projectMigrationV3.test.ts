// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  err,
  FxError,
  Inputs,
  ok,
  Platform,
  Result,
  SettingsFileName,
  SettingsFolderName,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv from "mocked-env";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import * as yaml from "js-yaml";
import { getProjectMigratorMW } from "../../../src/core/middleware/projectMigrator";
import { MockTools, MockUserInteraction, randomAppName } from "../utils";
import { CoreHookContext } from "../../../src/core/types";
import { setTools } from "../../../src/core/globalVars";
import { MigrationContext } from "../../../src/core/middleware/utils/migrationContext";
import {
  generateAppYml,
  generateSettingsJson,
  replacePlaceholderForManifests,
  statesMigration,
  updateLaunchJson,
  migrate,
  wrapRunMigration,
  checkVersionForMigration,
  VersionState,
  configsMigration,
} from "../../../src/core/middleware/projectMigratorV3";
import * as MigratorV3 from "../../../src/core/middleware/projectMigratorV3";
import { getProjectVersion } from "../../../src/core/middleware/utils/v3MigrationUtils";

let mockedEnvRestore: () => void;

describe("ProjectMigratorMW", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, ".fx"));
    mockedEnvRestore = mockedEnv({
      TEAMSFX_V3_MIGRATION: "true",
      TEAMSFX_V3: "false",
    });
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
    mockedEnvRestore();
  });

  it("happy path", async () => {
    sandbox.stub(MockUserInteraction.prototype, "showMessage").resolves(ok("Upgrade"));
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [getProjectMigratorMW()],
    });

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isOk());
    } finally {
      await fs.rmdir(inputs.projectPath!, { recursive: true });
    }
  });

  it("user cancel", async () => {
    sandbox
      .stub(MockUserInteraction.prototype, "showMessage")
      .resolves(err(new Error("user cancel") as FxError));
    const tools = new MockTools();
    setTools(tools);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    class MyClass {
      tools = tools;
      async other(inputs: Inputs, ctx?: CoreHookContext): Promise<Result<any, FxError>> {
        return ok("");
      }
    }
    hooks(MyClass, {
      other: [getProjectMigratorMW()],
    });

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const my = new MyClass();
    try {
      const res = await my.other(inputs);
      assert.isTrue(res.isErr());
    } finally {
      await fs.rmdir(inputs.projectPath!, { recursive: true });
    }
  });

  it("wrap run error ", async () => {
    const tools = new MockTools();
    setTools(tools);
    sandbox.stub(MigratorV3, "migrate").throws(new Error("mocker error"));
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const ctx = {
      arguments: [inputs],
    };
    const context = await MigrationContext.create(ctx);
    const res = wrapRunMigration(context, migrate);
  });
});

describe("MigrationContext", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
    await fs.ensureDir(path.join(projectPath, ".fx"));
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
    mockedEnvRestore();
  });

  it("happy path", async () => {
    const tools = new MockTools();
    setTools(tools);

    const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
    inputs.projectPath = projectPath;
    const ctx = {
      arguments: [inputs],
    };
    const context = await MigrationContext.create(ctx);
    let res = await context.backup(".fx");
    assert.isTrue(res);
    res = await context.backup("no-exist");
    assert.isFalse(res);
    await context.fsWriteFile("a", "test-data");
    await context.fsCopy("a", "a-copy");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "a-copy")));
    await context.fsEnsureDir("b/c");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "b/c")));
    await context.fsCreateFile("d");
    assert.isTrue(await fs.pathExists(path.join(context.projectPath, "d")));
    const modifiedPaths = context.getModifiedPaths();
    assert.isTrue(modifiedPaths.includes("a"));
    assert.isTrue(modifiedPaths.includes("a-copy"));
    assert.isTrue(modifiedPaths.includes("b"));
    assert.isTrue(modifiedPaths.includes("b/c"));
    assert.isTrue(modifiedPaths.includes("d"));

    await context.cleanModifiedPaths();
    assert.isEmpty(context.getModifiedPaths());

    context.addReport("test report");
    context.addTelemetryProperties({ testProperrty: "test property" });
    await context.restoreBackup();
    await context.cleanTeamsfx();
  });
});

describe("generateSettingsJson", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const oldProjectSettings = await readOldProjectSettings(projectPath);

    await generateSettingsJson(migrationContext);

    assert.isTrue(
      await fs.pathExists(path.join(projectPath, SettingsFolderName, SettingsFileName))
    );
    const newSettings = await readSettingJson(projectPath);
    assert.equal(newSettings.trackingId, oldProjectSettings.projectId);
    assert.equal(newSettings.version, "3.0.0");
  });

  it("no project id", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    delete projectSetting.projectId;
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateSettingsJson(migrationContext);

    const newSettings = await readSettingJson(projectPath);
    assert.isTrue(newSettings.hasOwnProperty("trackingId")); // will auto generate a new trackingId if old project does not have project id
  });
});

describe("generateAppYml-js/ts", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success in happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);

    await generateAppYml(migrationContext);

    const appYamlPath = path.join(projectPath, Constants.appYmlPath);
    assert.isTrue(await fs.pathExists(appYamlPath));
    const appYaml: any = yaml.load(await fs.readFile(appYamlPath, "utf8"));
    // validate basic part
    assert.equal(appYaml.version, "1.0.0");
    assert.exists(getAction(appYaml.provision, "arm/deploy"));
    assert.exists(getAction(appYaml.registerApp, "teamsApp/create"));
    assert.exists(getAction(appYaml.configureApp, "teamsApp/validate"));
    assert.exists(getAction(appYaml.configureApp, "teamsApp/createAppPackage"));
    assert.exists(getAction(appYaml.configureApp, "teamsApp/update"));
    assert.exists(getAction(appYaml.publish, "teamsApp/validate"));
    assert.exists(getAction(appYaml.publish, "teamsApp/createAppPackage"));
    assert.exists(getAction(appYaml.publish, "teamsApp/publishAppPackage"));
    // validate AAD part
    assert.exists(getAction(appYaml.registerApp, "aadApp/create"));
    assert.exists(getAction(appYaml.configureApp, "aadApp/update"));
    // validate tab part
    const npmCommandActions: Array<any> = getAction(appYaml.deploy, "npm/command");
    assert.exists(
      npmCommandActions.find(
        (item) => item.with.workingDirectory === "tabs" && item.with.args === "install"
      )
    );
    assert.exists(
      npmCommandActions.find(
        (item) => item.with.workingDirectory === "tabs" && item.with.args === "run build"
      )
    );
    assert.exists(getAction(appYaml.deploy, "azureStorage/deploy"));
  });

  it("should not generate AAD part if AAD plugin not activated", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.solutionSettings.activeResourcePlugins = (<Array<string>>(
      projectSetting.solutionSettings.activeResourcePlugins
    )).filter((item) => item !== "fx-resource-aad-app-for-teams"); // remove AAD plugin
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    const appYaml: any = yaml.load(
      await fs.readFile(path.join(projectPath, Constants.appYmlPath), "utf8")
    );

    assert.isEmpty(getAction(appYaml.registerApp, "aadApp/create"));
    assert.isEmpty(getAction(appYaml.configureApp, "aadApp/update"));
  });

  it("should not generate tab part if frontend hosting plugin not activated", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const projectSetting = await readOldProjectSettings(projectPath);
    projectSetting.solutionSettings.activeResourcePlugins = (<Array<string>>(
      projectSetting.solutionSettings.activeResourcePlugins
    )).filter((item) => item !== "fx-resource-frontend-hosting"); // remove frontend hosting plugin
    await fs.writeJson(
      path.join(projectPath, Constants.oldProjectSettingsFilePath),
      projectSetting
    );

    await generateAppYml(migrationContext);

    const appYaml: any = yaml.load(
      await fs.readFile(path.join(projectPath, Constants.appYmlPath), "utf8")
    );

    assert.isEmpty(getAction(appYaml.provision, "azureStorage/enableStaticWebsite"));
    const npmCommandActions: Array<any> = getAction(appYaml.deploy, "npm/command");
    assert.isEmpty(npmCommandActions.filter((item) => item.with.workingDirectory === "tabs"));
    assert.isEmpty(getAction(appYaml.deploy, "azureStorage/deploy"));
  });
});

describe("replacePlaceholderForManifests", () => {
  const sandbox = sinon.createSandbox();
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("happy path: aad manifest exists", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);

    // Action
    await replacePlaceholderForManifests(migrationContext);

    // Assert
    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const manifestPath = path.join(appPackageFolderPath, "manifest.template.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.template.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(manifest, manifestExpeceted);

    const aadManifestPath = path.join(projectPath, "aad.manifest.template.json");
    assert.isTrue(await fs.pathExists(aadManifestPath));
    const aadManifest = (await fs.readFile(aadManifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const aadManifestExpected = (
      await fs.readFile(path.join(projectPath, "expected", "aad.manifest.template.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(aadManifest, aadManifestExpected);
  });

  it("happy path: aad manifest does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    await fs.remove(path.join(projectPath, "templates/appPackage/aad.template.json"));

    // Action
    await replacePlaceholderForManifests(migrationContext);

    // Assert
    const appPackageFolderPath = path.join(projectPath, "appPackage");
    assert.isTrue(await fs.pathExists(appPackageFolderPath));

    const resourcesPath = path.join(appPackageFolderPath, "resources", "test.png");
    assert.isTrue(await fs.pathExists(resourcesPath));

    const manifestPath = path.join(appPackageFolderPath, "manifest.template.json");
    assert.isTrue(await fs.pathExists(manifestPath));
    const manifest = (await fs.readFile(manifestPath, "utf-8"))
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    const manifestExpeceted = (
      await fs.readFile(path.join(projectPath, "expected", "manifest.template.json"), "utf-8")
    )
      .replace(/\s/g, "")
      .replace(/\t/g, "")
      .replace(/\n/g, "");
    assert.equal(manifest, manifestExpeceted);

    const aadManifestPath = path.join(projectPath, "aad.manifest.template.json");
    assert.isFalse(await fs.pathExists(aadManifestPath));
  });

  it("migrate manifests failed: appPackage does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(false);

    try {
      await replacePlaceholderForManifests(migrationContext);
    } catch (error) {
      assert.equal(error.name, "ReadFileError");
      assert.equal(error.innerError.message, "templates/appPackage does not exist");
    }
  });

  it("migrate manifests failed: provision.bicep does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await fs.ensureDir(path.join(projectPath, "appPackage"));

    try {
      await replacePlaceholderForManifests(migrationContext);
    } catch (error) {
      assert.equal(error.name, "ReadFileError");
      assert.equal(error.innerError.message, "templates/azure/provision.bicep does not exist");
    }
  });

  it("migrate manifests failed: teams app manifest does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    // Stub
    sandbox.stub(migrationContext, "backup").resolves(true);
    await copyTestProject(Constants.manifestsMigrationHappyPath, projectPath);
    await fs.remove(path.join(projectPath, "templates/appPackage/manifest.template.json"));

    try {
      await replacePlaceholderForManifests(migrationContext);
    } catch (error) {
      assert.equal(error.name, "ReadFileError");
      assert.equal(
        error.innerError.message,
        "templates/appPackage/manifest.template.json does not exist"
      );
    }
  });
});

describe("updateLaunchJson", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("should success in happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);

    await updateLaunchJson(migrationContext);

    assert.isTrue(
      await fs.pathExists(path.join(projectPath, "teamsfx/backup/.vscode/launch.json"))
    );
    const updatedLaunchJson = await fs.readJson(path.join(projectPath, Constants.launchJsonPath));
    assert.equal(
      updatedLaunchJson.configurations[0].url,
      "https://teams.microsoft.com/l/app/${dev:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[1].url,
      "https://teams.microsoft.com/l/app/${dev:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[2].url,
      "https://teams.microsoft.com/l/app/${local:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[3].url,
      "https://teams.microsoft.com/l/app/${local:teamsAppId}?installAppPackage=true&webjoin=true&${account-hint}"
    );
    assert.equal(
      updatedLaunchJson.configurations[4].url,
      "https://outlook.office.com/host/${local:teamsAppInternalId}?${account-hint}" // for M365 app
    );
    assert.equal(
      updatedLaunchJson.configurations[5].url,
      "https://outlook.office.com/host/${local:teamsAppInternalId}?${account-hint}" // for M365 app
    );
  });
});

describe("stateMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await statesMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "dev"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.dev")));
    const testEnvContent_dev = await readEnvFile(path.join(projectPath, "teamsfx"), "dev");
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "local"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.local")));
    const testEnvContent_local = await readEnvFile(path.join(projectPath, "teamsfx"), "local");
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });

  it("ReadFileError: .fx/states does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject("happyPathEmpty", projectPath);
    try {
      await statesMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, "ReadFileError");
      assert.equal(error.innerError.message, ".fx/states does not exist");
    }
  });
});

describe("configMigration", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
  });

  it("happy path", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject(Constants.happyPathTestProject, projectPath);
    await configsMigration(migrationContext);

    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx")));

    const trueEnvContent_dev = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "config.dev"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.dev")));
    const testEnvContent_dev = await readEnvFile(path.join(projectPath, "teamsfx"), "dev");
    assert.equal(testEnvContent_dev, trueEnvContent_dev);

    const trueEnvContent_local = await readEnvFile(
      getTestAssetsPath(path.join(Constants.happyPathTestProject, "testCaseFiles")),
      "config.local"
    );
    assert.isTrue(await fs.pathExists(path.join(projectPath, "teamsfx", ".env.local")));
    const testEnvContent_local = await readEnvFile(path.join(projectPath, "teamsfx"), "local");
    assert.equal(testEnvContent_local, trueEnvContent_local);
  });

  it("ReadFileError: .fx/configs does not exist", async () => {
    const migrationContext = await mockMigrationContext(projectPath);

    await copyTestProject("happyPathEmpty", projectPath);
    try {
      await configsMigration(migrationContext);
    } catch (error) {
      assert.equal(error.name, "ReadFileError");
      assert.equal(error.innerError.message, ".fx/configs does not exist");
    }
  });
});

describe("Migration utils", () => {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    await fs.ensureDir(projectPath);
  });

  afterEach(async () => {
    await fs.remove(projectPath);
    sandbox.restore();
  });

  it("checkVersionForMigration V2", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state, VersionState.upgradeable);
  });

  it("checkVersionForMigration V3", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readJson").resolves("3.0.0");
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state, VersionState.compatible);
  });

  it("checkVersionForMigration empty", async () => {
    const migrationContext = await mockMigrationContext(projectPath);
    await copyTestProject(Constants.happyPathTestProject, projectPath);
    sandbox.stub(fs, "pathExists").resolves(false);
    const state = await checkVersionForMigration(migrationContext);
    assert.equal(state, VersionState.unsupported);
  });
});

async function mockMigrationContext(projectPath: string): Promise<MigrationContext> {
  const inputs: Inputs = { platform: Platform.VSCode, ignoreEnvInfo: true };
  inputs.projectPath = projectPath;
  const ctx = {
    arguments: [inputs],
  };
  return await MigrationContext.create(ctx);
}

function getTestAssetsPath(projectName: string): string {
  return path.join("tests/core/middleware/testAssets/v3Migration", projectName.toString());
}

async function copyTestProject(projectName: string, targetPath: string): Promise<void> {
  await fs.copy(getTestAssetsPath(projectName), targetPath);
}

async function readOldProjectSettings(projectPath: string): Promise<any> {
  return await fs.readJson(path.join(projectPath, Constants.oldProjectSettingsFilePath));
}

async function readSettingJson(projectPath: string): Promise<any> {
  return await fs.readJson(path.join(projectPath, Constants.settingsFilePath));
}

async function readEnvFile(projectPath: string, env: string): Promise<any> {
  return await fs.readFileSync(path.join(projectPath, ".env." + env)).toString();
}

function getAction(lifecycleDefinition: Array<any>, actionName: string): any[] {
  if (lifecycleDefinition) {
    return lifecycleDefinition.filter((item) => item.uses === actionName);
  }
  return [];
}

const Constants = {
  happyPathTestProject: "happyPath",
  settingsFilePath: "teamsfx/settings.json",
  oldProjectSettingsFilePath: ".fx/configs/projectSettings.json",
  appYmlPath: "teamsfx/app.yml",
  manifestsMigrationHappyPath: "manifestsHappyPath",
  launchJsonPath: ".vscode/launch.json",
};
