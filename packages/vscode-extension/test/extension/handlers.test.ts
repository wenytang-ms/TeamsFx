import * as chai from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import * as sinon from "sinon";
import { stubInterface } from "ts-sinon";
import * as util from "util";
import * as uuid from "uuid";
import * as vscode from "vscode";

import {
  ConfigFolderName,
  err,
  FxError,
  Inputs,
  IProgressHandler,
  ok,
  Platform,
  ProjectSettings,
  ProjectSettingsFileName,
  Result,
  StaticOptions,
  Stage,
  UserError,
  Void,
  VsCodeEnv,
  PathNotExistError,
} from "@microsoft/teamsfx-api";
import { DepsManager, DepsType } from "@microsoft/teamsfx-core/build/common/deps-checker";
import * as globalState from "@microsoft/teamsfx-core/build/common/globalState";
import { CollaborationState } from "@microsoft/teamsfx-core/build/common/permissionInterface";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import { CoreHookContext } from "@microsoft/teamsfx-core/build/core/types";

import * as StringResources from "../../package.nls.json";
import { AzureAccountManager } from "../../src/commonlib/azureLogin";
import M365TokenInstance from "../../src/commonlib/m365Login";
import { SUPPORTED_SPFX_VERSION } from "../../src/constants";
import { PanelType } from "../../src/controls/PanelType";
import { WebviewPanel } from "../../src/controls/webviewPanel";
import * as debugCommonUtils from "../../src/debug/commonUtils";
import * as teamsAppInstallation from "../../src/debug/teamsAppInstallation";
import { vscodeHelper } from "../../src/debug/depsChecker/vscodeHelper";
import * as debugProvider from "../../src/debug/teamsfxDebugProvider";
import * as taskHandler from "../../src/debug/teamsfxTaskHandler";
import { ExtensionErrors } from "../../src/error";
import * as extension from "../../src/extension";
import * as globalVariables from "../../src/globalVariables";
import * as handlers from "../../src/handlers";
import { VsCodeUI } from "../../src/qm/vsc_ui";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as extTelemetryEvents from "../../src/telemetry/extTelemetryEvents";
import accountTreeViewProviderInstance from "../../src/treeview/account/accountTreeViewProvider";
import envTreeProviderInstance from "../../src/treeview/environmentTreeViewProvider";
import TreeViewManagerInstance from "../../src/treeview/treeViewManager";
import * as commonUtils from "../../src/utils/commonUtils";
import * as localizeUtils from "../../src/utils/localizeUtils";
import { MockCore } from "../mocks/mockCore";
import * as commonTools from "@microsoft/teamsfx-core/build/common/tools";
import { VsCodeLogProvider } from "../../src/commonlib/log";
import { ProgressHandler } from "../../src/progressHandler";
import { TreatmentVariableValue } from "../../src/exp/treatmentVariables";
import { assert } from "console";
import { AppStudioClient } from "@microsoft/teamsfx-core/build/component/resource/appManifest/appStudioClient";
import { AppDefinition } from "@microsoft/teamsfx-core/build/component/resource/appManifest/interfaces/appDefinition";

describe("handlers", () => {
  describe("activate()", function () {
    const sandbox = sinon.createSandbox();
    let setStatusChangeMap: any;

    beforeEach(() => {
      sandbox.stub(accountTreeViewProviderInstance, "subscribeToStatusChanges");
      sandbox.stub(vscode.extensions, "getExtension").returns(undefined);
      sandbox.stub(TreeViewManagerInstance, "getTreeView").returns(undefined);
      sandbox.stub(ExtTelemetry, "dispose");
    });

    afterEach(() => {
      sandbox.restore();
    });

    it("No globalState error", async () => {
      const result = await handlers.activate();
      chai.assert.deepEqual(result.isOk() ? result.value : result.error.name, {});
    });

    it("Valid project", async () => {
      sandbox.stub(projectSettingsHelper, "isValidProject").returns(true);
      const sendTelemetryStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      const addSharedPropertyStub = sandbox.stub(ExtTelemetry, "addSharedProperty");
      const result = await handlers.activate();

      chai.assert.isTrue(addSharedPropertyStub.called);
      chai.assert.isTrue(sendTelemetryStub.calledOnceWith("open-teams-app"));
      chai.assert.deepEqual(result.isOk() ? result.value : result.error.name, {});
    });
  });

  it("getSystemInputs()", () => {
    sinon.stub(vscodeHelper, "checkerEnabled").returns(false);
    const input: Inputs = handlers.getSystemInputs();

    chai.expect(input.platform).equals(Platform.VSCode);
  });

  it("getAzureProjectConfigV3", async () => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(handlers, "core").value(new MockCore());
    sandbox.stub(handlers, "getSystemInputs").returns({} as Inputs);
    const fake_config_v3 = {
      projectSettings: {
        appName: "fake_test",
        projectId: "fake_projectId",
      },
      envInfos: {},
    };
    sandbox.stub(MockCore.prototype, "getProjectConfigV3").resolves(ok(fake_config_v3));
    const res = await handlers.getAzureProjectConfigV3();
    chai.assert.exists(res?.projectSettings);
    chai.assert.equal(res?.projectSettings.appName, "fake_test");
    chai.assert.equal(res?.projectSettings.projectId, "fake_projectId");
    sandbox.restore();
  });

  it("getAzureProjectConfigV3 return undefined", async () => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(handlers, "core").value(new MockCore());
    sandbox.stub(handlers, "getSystemInputs").returns({} as Inputs);
    sandbox
      .stub(MockCore.prototype, "getProjectConfigV3")
      .resolves(err(new PathNotExistError("path not exist", "fake path")));
    const res = await handlers.getAzureProjectConfigV3();
    chai.assert.isUndefined(res);
    sandbox.restore();
  });

  it("getSettingsVersion in v3", async () => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(commonTools, "isV3Enabled").returns(true);
    sandbox.stub(handlers, "core").value(new MockCore());
    sandbox.stub(handlers, "getSystemInputs").returns({} as Inputs);
    sandbox
      .stub(MockCore.prototype, "getSettings")
      .resolves(ok({ version: "3.0.0" } as ProjectSettings));
    const res = await handlers.getSettingsVersion();
    chai.assert.equal(res, "3.0.0");
    sandbox.restore();
  });

  it("openBackupConfigMd", async () => {
    const workspacePath = "test";
    const filePath = path.join(workspacePath, ".backup", "backup-config-change-logs.md");

    const openTextDocument = sinon.stub(vscode.workspace, "openTextDocument").resolves();
    const executeCommand = sinon.stub(vscode.commands, "executeCommand").resolves();

    await handlers.openBackupConfigMd(workspacePath, filePath);

    chai.assert.isTrue(openTextDocument.calledOnce);
    chai.assert.isTrue(
      executeCommand.calledOnceWithExactly("markdown.showPreview", vscode.Uri.file(filePath))
    );
    openTextDocument.restore();
    executeCommand.restore();
  });

  it("addFileSystemWatcher", async () => {
    const workspacePath = "test";

    const watcher = { onDidCreate: () => ({ dispose: () => undefined }) } as any;
    const createWatcher = sinon.stub(vscode.workspace, "createFileSystemWatcher").returns(watcher);
    const listener = sinon.stub(watcher, "onDidCreate").resolves();

    handlers.addFileSystemWatcher(workspacePath);

    chai.assert.isTrue(createWatcher.calledTwice);
    chai.assert.isTrue(listener.calledTwice);
    createWatcher.restore();
    listener.restore();
  });

  describe("command handlers", function () {
    this.afterEach(() => {
      sinon.restore();
    });

    it("createNewProjectHandler()", async () => {
      const clock = sinon.useFakeTimers();

      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(commonUtils, "isExistingTabApp").returns(Promise.resolve(false));
      const sendTelemetryEventFunc = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createProject = sinon.spy(handlers.core, "createProject");
      const executeCommandFunc = sinon.stub(vscode.commands, "executeCommand");
      const globalStateUpdateStub = sinon.stub(globalState, "globalStateUpdate");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.createNewProjectHandler();

      chai.assert.isTrue(
        sendTelemetryEventFunc.calledWith(extTelemetryEvents.TelemetryEvent.CreateProjectStart)
      );
      chai.assert.isTrue(
        sendTelemetryEventFunc.calledWith(extTelemetryEvents.TelemetryEvent.CreateProject)
      );
      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(executeCommandFunc.calledOnceWith("vscode.openFolder"));
      sinon.restore();
      clock.restore();
    });

    it("provisionHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const provisionResources = sinon.spy(handlers.core, "provisionResources");
      sinon.stub(envTreeProviderInstance, "reloadEnvironments");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.provisionHandler();

      sinon.assert.calledOnce(provisionResources);
      sinon.restore();
    });

    it("deployHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const deployArtifacts = sinon.spy(handlers.core, "deployArtifacts");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.deployHandler();

      sinon.assert.calledOnce(deployArtifacts);
      sinon.restore();
    });

    it("publishHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const publishApplication = sinon.spy(handlers.core, "publishApplication");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.publishHandler();

      sinon.assert.calledOnce(publishApplication);
      sinon.restore();
    });

    it("buildPackageHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      await handlers.buildPackageHandler();

      // should show error for invalid project
      sinon.assert.calledOnce(sendTelemetryErrorEvent);
      sinon.restore();
    });

    it("validateManifestHandler()", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      await handlers.validateManifestHandler();

      sinon.assert.calledOnce(sendTelemetryErrorEvent);
      sinon.restore();
    });

    it("validateManifestHandler() - V3", async () => {
      sinon.stub(commonTools, "isV3Enabled").returns(true);
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(localizeUtils, "localize").returns("");

      const res = await handlers.validateManifestHandler();

      chai.assert(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, ExtensionErrors.DefaultManifestTemplateNotExistsError);
      }
      sinon.restore();
    });

    it("debugHandler()", async () => {
      const sendTelemetryEventStub = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const executeCommandStub = sinon.stub(vscode.commands, "executeCommand");

      await handlers.debugHandler();

      sinon.assert.calledOnceWithExactly(executeCommandStub, "workbench.action.debug.start");
      sinon.assert.calledOnce(sendTelemetryEventStub);
      sinon.restore();
    });

    it("treeViewPreviewHandler()", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(debugCommonUtils, "getDebugConfig").resolves({ appId: "appId" });
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      let ignoreEnvInfo: boolean | undefined = undefined;
      let localDebugCalled = 0;
      sinon
        .stub(handlers.core, "localDebug")
        .callsFake(
          async (
            inputs: Inputs,
            ctx?: CoreHookContext | undefined
          ): Promise<Result<Void, FxError>> => {
            ignoreEnvInfo = inputs.ignoreEnvInfo;
            localDebugCalled += 1;
            return ok({});
          }
        );
      const mockProgressHandler = stubInterface<IProgressHandler>();
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      sinon.stub(VsCodeUI.prototype, "createProgressBar").returns(mockProgressHandler);
      sinon.stub(VsCodeUI.prototype, "openUrl");
      sinon.stub(debugProvider, "generateAccountHint");

      const result = await handlers.treeViewPreviewHandler("local");

      chai.assert.isTrue(result.isOk());
    });

    it("selectTutorialsHandler()", async () => {
      sinon.stub(localizeUtils, "localize").returns("");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(TreatmentVariableValue, "inProductDoc").value(true);
      let tutorialOptions: StaticOptions[] = [];
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: (options: any) => {
          tutorialOptions = options.options;
          return Promise.resolve(ok({ type: "success", result: { id: "test", data: "data" } }));
        },
        openUrl: () => Promise.resolve(ok(true)),
      });

      const result = await handlers.selectTutorialsHandler();

      chai.assert.equal(tutorialOptions.length, 6);
      chai.assert.isTrue(result.isOk());
    });
  });

  describe("runCommand()", function () {
    this.afterEach(() => {
      sinon.restore();
    });

    it("openConfigStateFile() - local", async () => {
      sinon.stub(localizeUtils, "localize").callsFake((key: string) => {
        return key;
      });

      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: ProjectSettings = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, ProjectSettingsFileName);
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await handlers.openConfigStateFile([{ type: "state" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.EnvStateNotFoundError);
        chai.assert.equal(
          res.error.message,
          util.format(localizeUtils.localize("teamstoolkit.handlers.localStateFileNotFound"), env)
        );
      }
    });

    it("openConfigStateFile() - env - FileNotFound", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: ProjectSettings = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, ProjectSettingsFileName);
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await handlers.openConfigStateFile([{ type: "env" }]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.EnvFileNotFoundError);
      }
    });

    it("openConfigStateFile() - InvalidArgs", async () => {
      const env = "local";
      const tmpDir = fs.mkdtempSync(path.resolve("./tmp"));

      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(tmpDir));
      const projectSettings: ProjectSettings = {
        appName: "myapp",
        version: "1.0.0",
        projectId: "123",
      };
      const configFolder = path.resolve(tmpDir, `.${ConfigFolderName}`, "configs");
      await fs.mkdir(configFolder, { recursive: true });
      const settingsFile = path.resolve(configFolder, ProjectSettingsFileName);
      await fs.writeJSON(settingsFile, JSON.stringify(projectSettings, null, 4));

      sinon.stub(globalVariables, "context").value({ extensionPath: path.resolve("../../") });
      sinon.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: env })),
      });

      const res = await handlers.openConfigStateFile([]);
      await fs.remove(tmpDir);

      if (res) {
        chai.assert.isTrue(res.isErr());
        chai.assert.equal(res.error.name, ExtensionErrors.InvalidArgs);
      }
    });

    it("create sample with projectid", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createProject = sinon.spy(handlers.core, "createProject");
      sinon.stub(vscode.commands, "executeCommand");
      const inputs = { projectId: uuid.v4(), platform: Platform.VSCode };
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.runCommand(Stage.create, inputs);

      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(createProject.args[0][0].projectId != undefined);
      chai.assert.isTrue(sendTelemetryEvent.args[0][1]!["new-project-id"] != undefined);
    });

    it("create from scratch without projectid", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createProject = sinon.spy(handlers.core, "createProject");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(true);

      await handlers.runCommand(Stage.create);

      sinon.restore();
      sinon.assert.calledOnce(createProject);
      chai.assert.isTrue(createProject.args[0][0].projectId != undefined);
      chai.assert.isTrue(sendTelemetryEvent.args[0][1]!["new-project-id"] != undefined);
    });

    it("provisionResources", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const provisionResources = sinon.spy(handlers.core, "provisionResources");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.runCommand(Stage.provision);

      sinon.restore();
      sinon.assert.calledOnce(provisionResources);
    });

    it("deployArtifacts", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const deployArtifacts = sinon.spy(handlers.core, "deployArtifacts");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.runCommand(Stage.deploy);

      sinon.restore();
      sinon.assert.calledOnce(deployArtifacts);
    });

    it("deployAadManifest", async () => {
      const sandbox = sinon.createSandbox();
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const deployAadManifest = sandbox.spy(handlers.core, "deployAadManifest");
      sandbox.stub(vscodeHelper, "checkerEnabled").returns(false);
      const input: Inputs = handlers.getSystemInputs();
      await handlers.runCommand(Stage.deployAad, input);

      sandbox.assert.calledOnce(deployAadManifest);
      sandbox.restore();
    });

    it("deployAadManifest happy path", async () => {
      const sandbox = sinon.createSandbox();
      sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
      sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sandbox.stub(handlers.core, "deployAadManifest").resolves(ok("test_success"));
      sandbox.stub(vscodeHelper, "checkerEnabled").returns(false);
      const input: Inputs = handlers.getSystemInputs();
      const res = await handlers.runCommand(Stage.deployAad, input);
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.strictEqual(res.value, "test_success");
      }
      sandbox.restore();
    });

    it("localDebug", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      let ignoreEnvInfo: boolean | undefined = undefined;
      let localDebugCalled = 0;
      sinon
        .stub(handlers.core, "localDebug")
        .callsFake(
          async (
            inputs: Inputs,
            ctx?: CoreHookContext | undefined
          ): Promise<Result<Void, FxError>> => {
            ignoreEnvInfo = inputs.ignoreEnvInfo;
            localDebugCalled += 1;
            return ok({});
          }
        );

      await handlers.runCommand(Stage.debug);

      sinon.restore();
      chai.expect(ignoreEnvInfo).to.equal(false);
      chai.expect(localDebugCalled).equals(1);
    });

    it("publishApplication", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const publishApplication = sinon.spy(handlers.core, "publishApplication");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.runCommand(Stage.publish);

      sinon.restore();
      sinon.assert.calledOnce(publishApplication);
    });

    it("createEnv", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const createEnv = sinon.spy(handlers.core, "createEnv");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      await handlers.runCommand(Stage.createEnv);

      sinon.restore();
      sinon.assert.calledOnce(createEnv);
    });
  });

  describe("detectVsCodeEnv()", function () {
    this.afterEach(() => {
      sinon.restore();
    });

    it("locally run", () => {
      const expectedResult = {
        extensionKind: vscode.ExtensionKind.UI,
        id: "",
        extensionUri: vscode.Uri.file(""),
        extensionPath: "",
        isActive: true,
        packageJSON: {},
        exports: undefined,
        activate: sinon.spy(),
      };
      const getExtension = sinon
        .stub(vscode.extensions, "getExtension")
        .callsFake((name: string) => {
          return expectedResult;
        });

      chai.expect(handlers.detectVsCodeEnv()).equals(VsCodeEnv.local);
      getExtension.restore();
    });

    it("Remotely run", () => {
      const expectedResult = {
        extensionKind: vscode.ExtensionKind.Workspace,
        id: "",
        extensionUri: vscode.Uri.file(""),
        extensionPath: "",
        isActive: true,
        packageJSON: {},
        exports: undefined,
        activate: sinon.spy(),
      };
      const getExtension = sinon
        .stub(vscode.extensions, "getExtension")
        .callsFake((name: string) => {
          return expectedResult;
        });

      chai
        .expect(handlers.detectVsCodeEnv())
        .oneOf([VsCodeEnv.remote, VsCodeEnv.codespaceVsCode, VsCodeEnv.codespaceBrowser]);
      getExtension.restore();
    });
  });

  it("openWelcomeHandler", async () => {
    const executeCommands = sinon.stub(vscode.commands, "executeCommand");
    const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.openWelcomeHandler();

    sinon.assert.calledOnceWithExactly(
      executeCommands,
      "workbench.action.openWalkthrough",
      "TeamsDevApp.ms-teams-vscode-extension#teamsToolkitGetStarted"
    );
    executeCommands.restore();
    sendTelemetryEvent.restore();
  });

  it("openSamplesHandler", async () => {
    const createOrShow = sinon.stub(WebviewPanel, "createOrShow");
    const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.openSamplesHandler();

    sinon.assert.calledOnceWithExactly(createOrShow, PanelType.SampleGallery, false);
    createOrShow.restore();
    sendTelemetryEvent.restore();
  });

  it("signOutM365", async () => {
    const signOut = sinon.stub(M365TokenInstance, "signout");
    const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
    sinon.stub(envTreeProviderInstance, "reloadEnvironments");

    await handlers.signOutM365(false);

    sinon.assert.calledOnce(signOut);
    signOut.restore();
    sendTelemetryEvent.restore();
  });

  it("signOutAzure", async () => {
    Object.setPrototypeOf(AzureAccountManager, sinon.stub());
    const signOut = sinon.stub(AzureAccountManager.getInstance(), "signout");
    const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");

    await handlers.signOutAzure(false);

    sinon.assert.calledOnce(signOut);
    signOut.restore();
    sendTelemetryEvent.restore();
  });

  describe("decryptSecret", function () {
    this.afterEach(() => {
      sinon.restore();
    });
    it("successfully update secret", async () => {
      sinon.stub(globalVariables, "context").value({ extensionPath: "" });
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const decrypt = sinon.spy(handlers.core, "decrypt");
      const encrypt = sinon.spy(handlers.core, "encrypt");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(true);
      const editBuilder = sinon.spy();
      sinon.stub(vscode.window, "activeTextEditor").value({
        edit: function (callback: (eb: any) => void) {
          callback({
            replace: editBuilder,
          });
        },
      });
      sinon.stub(extension, "VS_CODE_UI").value({
        inputText: () => Promise.resolve(ok({ type: "success", result: "inputValue" })),
      });
      const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

      await handlers.decryptSecret("test", range);

      sinon.assert.calledOnce(decrypt);
      sinon.assert.calledOnce(encrypt);
      sinon.assert.calledOnce(editBuilder);
      sinon.assert.calledTwice(sendTelemetryEvent);
      sinon.assert.notCalled(sendTelemetryErrorEvent);
      sinon.restore();
    });

    it("failed to update due to corrupted secret", async () => {
      sinon.stub(globalVariables, "context").value({ extensionPath: "" });
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      const decrypt = sinon.stub(handlers.core, "decrypt");
      decrypt.returns(Promise.resolve(err(new UserError("", "fake error", ""))));
      const encrypt = sinon.spy(handlers.core, "encrypt");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(true);
      const editBuilder = sinon.spy();
      sinon.stub(vscode.window, "activeTextEditor").value({
        edit: function (callback: (eb: any) => void) {
          callback({
            replace: editBuilder,
          });
        },
      });
      const showMessage = sinon.stub(vscode.window, "showErrorMessage");
      const range = new vscode.Range(new vscode.Position(0, 10), new vscode.Position(0, 15));

      await handlers.decryptSecret("test", range);

      sinon.assert.calledOnce(decrypt);
      sinon.assert.notCalled(encrypt);
      sinon.assert.notCalled(editBuilder);
      sinon.assert.calledOnce(showMessage);
      sinon.assert.calledOnce(sendTelemetryEvent);
      sinon.assert.calledOnce(sendTelemetryErrorEvent);
      sinon.restore();
    });
  });

  describe("permissions", async function () {
    this.afterEach(() => {
      sinon.restore();
    });
    it("grant permission", async () => {
      sinon.restore();
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(commonUtils, "getProvisionSucceedFromEnv").resolves(true);
      sinon.stub(M365TokenInstance, "getJsonObject").resolves(
        ok({
          tid: "fake-tenant-id",
        })
      );

      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.parse("file://fakeProjectPath"));
      sinon.stub(globalVariables, "isSPFxProject").value(false);
      sinon.stub(commonUtils, "getM365TenantFromEnv").callsFake(async (env: string) => {
        return "fake-tenant-id";
      });

      sinon.stub(MockCore.prototype, "grantPermission").returns(
        Promise.resolve(
          ok({
            state: CollaborationState.OK,
            userInfo: {
              userObjectId: "fake-user-object-id",
              userPrincipalName: "fake-user-principle-name",
            },
            permissions: [
              {
                name: "name",
                type: "type",
                resourceId: "id",
                roles: ["Owner"],
              },
            ],
          })
        )
      );
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      const result = await handlers.grantPermission("env");
      chai.expect(result.isOk()).equals(true);
    });

    it("grant permission with empty tenant id", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(commonUtils, "getProvisionSucceedFromEnv").resolves(true);
      sinon.stub(M365TokenInstance, "getJsonObject").resolves(
        ok({
          tid: "fake-tenant-id",
        })
      );
      sinon.stub(commonUtils, "getM365TenantFromEnv").callsFake(async (env: string) => {
        return "";
      });

      const result = await handlers.grantPermission("env");

      if (result.isErr()) {
        throw new Error("Unexpected error: " + result.error.message);
      }

      chai.expect(result.isOk()).equals(true);
      chai.expect(result.value.state === CollaborationState.EmptyM365Tenant);
    });

    it("list collaborators", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(commonUtils, "getProvisionSucceedFromEnv").resolves(true);
      sinon.stub(M365TokenInstance, "getJsonObject").resolves(
        ok({
          tid: "fake-tenant-id",
        })
      );
      sinon.stub(commonUtils, "getM365TenantFromEnv").callsFake(async (env: string) => {
        return "fake-tenant-id";
      });

      await handlers.listCollaborator("env");
    });

    it("list collaborators with empty tenant id", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const sendTelemetryEvent = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const sendTelemetryErrorEvent = sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(commonUtils, "getProvisionSucceedFromEnv").resolves(true);
      sinon.stub(M365TokenInstance, "getJsonObject").resolves(
        ok({
          tid: "fake-tenant-id",
        })
      );
      sinon.stub(commonUtils, "getM365TenantFromEnv").callsFake(async (env: string) => {
        return "";
      });

      const showWarningMessage = sinon
        .stub(vscode.window, "showWarningMessage")
        .callsFake((message: string): any => {
          chai
            .expect(message)
            .equal(StringResources["teamstoolkit.commandsTreeViewProvider.emptyM365Tenant"]);
        });
      await handlers.listCollaborator("env");

      chai.expect(showWarningMessage.callCount).to.be.equal(1);
    });
  });

  describe("permission v3", function () {
    const sandbox = sinon.createSandbox();

    this.beforeEach(() => {
      sandbox.stub(commonTools, "isV3Enabled").returns(true);
    });

    this.afterEach(() => {
      sandbox.restore();
    });

    it("happy path: grant permission", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: "grantPermission" })),
      });
      sandbox.stub(MockCore.prototype, "grantPermission").returns(
        Promise.resolve(
          ok({
            state: CollaborationState.OK,
            userInfo: {
              userObjectId: "fake-user-object-id",
              userPrincipalName: "fake-user-principle-name",
            },
            permissions: [
              {
                name: "name",
                type: "type",
                resourceId: "id",
                roles: ["Owner"],
              },
            ],
          })
        )
      );
      sandbox.stub(vscodeHelper, "checkerEnabled").returns(false);

      const result = await handlers.manageCollaboratorHandler();
      chai.expect(result.isOk()).equals(true);
    });

    it("happy path: list collaborator", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () => Promise.resolve(ok({ type: "success", result: "listCollaborator" })),
      });
      sandbox.stub(MockCore.prototype, "listCollaborator").returns(
        Promise.resolve(
          ok({
            state: CollaborationState.OK,
            collaborators: [
              {
                userPrincipalName: "userPrincipalName",
                userObjectId: "userObjectId",
                isAadOwner: true,
                teamsAppResourceId: "teamsAppResourceId",
              },
            ],
          })
        )
      );
      sandbox.stub(vscodeHelper, "checkerEnabled").returns(false);
      const vscodeLogProviderInstance = VsCodeLogProvider.getInstance();
      sandbox.stub(vscodeLogProviderInstance, "outputChannel").value({
        name: "name",
        append: (value: string) => {},
        appendLine: (value: string) => {},
        replace: (value: string) => {},
        clear: () => {},
        show: (...params: any[]) => {},
        hide: () => {},
        dispose: () => {},
      });

      const result = await handlers.manageCollaboratorHandler();
      chai.expect(result.isOk()).equals(true);
    });

    it("User Cancel", async () => {
      sandbox.stub(handlers, "core").value(new MockCore());
      sandbox.stub(extension, "VS_CODE_UI").value({
        selectOption: () =>
          Promise.resolve(err(new UserError("source", "errorName", "errorMessage"))),
      });

      const result = await handlers.manageCollaboratorHandler();
      chai.expect(result.isErr()).equals(true);
    });
  });

  describe("manifest", () => {
    it("edit manifest template: local", async () => {
      sinon.restore();
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const openTextDocument = sinon
        .stub(vscode.workspace, "openTextDocument")
        .returns(new Promise<vscode.TextDocument>((resolve) => {}));
      sinon
        .stub(vscode.workspace, "workspaceFolders")
        .returns([{ uri: { fsPath: "c:\\manifestTestFolder" } }]);

      const args = [{ fsPath: "c:\\testPath\\manifest.local.json" }, "CodeLens"];
      await handlers.editManifestTemplate(args);
      chai.assert.isTrue(
        openTextDocument.calledOnceWith(
          "undefined/templates/appPackage/manifest.template.json" as any
        )
      );
    });

    it("edit manifest template: remote", async () => {
      sinon.restore();
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      const openTextDocument = sinon
        .stub(vscode.workspace, "openTextDocument")
        .returns(new Promise<vscode.TextDocument>((resolve) => {}));
      sinon
        .stub(vscode.workspace, "workspaceFolders")
        .returns([{ uri: { fsPath: "c:\\manifestTestFolder" } }]);

      const args = [{ fsPath: "c:\\testPath\\manifest.dev.json" }, "CodeLens"];
      await handlers.editManifestTemplate(args);
      chai.assert.isTrue(
        openTextDocument.calledOnceWith(
          "undefined/templates/appPackage/manifest.template.json" as any
        )
      );
    });
  });

  it("downloadSample", async () => {
    const inputs: Inputs = {
      scratch: "no",
      platform: Platform.VSCode,
    };
    sinon.stub(handlers, "core").value(new MockCore());
    const createProject = sinon.spy(handlers.core, "createProject");

    await handlers.downloadSample(inputs);

    inputs.stage = Stage.create;
    chai.assert.isTrue(createProject.calledOnceWith(inputs));
  });

  it("deployAadAppManifest", async () => {
    sinon.stub(handlers, "core").value(new MockCore());
    sinon.stub(ExtTelemetry, "sendTelemetryEvent");
    sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const deployArtifacts = sinon.spy(handlers.core, "deployArtifacts");
    await handlers.updateAadAppManifest([{ fsPath: "path/aad.dev.template" }, "CodeLens"]);
    sinon.assert.calledOnce(deployArtifacts);
    chai.assert.equal(deployArtifacts.getCall(0).args[0]["include-aad-manifest"], "yes");
    sinon.restore();
  });

  it("deployAadAppManifest v3", async () => {
    sinon.stub(commonTools, "isV3Enabled").returns(true);
    sinon.stub(vscodeHelper, "checkerEnabled").returns(false);
    sinon.stub(handlers, "core").value(new MockCore());
    sinon.stub(ExtTelemetry, "sendTelemetryEvent");
    sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const deployAadManifest = sinon.spy(handlers.core, "deployAadManifest");
    await handlers.updateAadAppManifest([{ fsPath: "path/aad.dev.template" }, "CodeLens"]);
    sinon.assert.calledOnce(deployAadManifest);
    chai.assert.equal(deployAadManifest.getCall(0).args[0]["include-aad-manifest"], "yes");
    deployAadManifest.restore();
    sinon.restore();
  });

  it("showError", async () => {
    sinon.stub(localizeUtils, "localize").returns("");
    const showErrorMessageStub = sinon
      .stub(vscode.window, "showErrorMessage")
      .callsFake((title: string, button: any) => {
        return Promise.resolve(button);
      });
    const sendTelemetryEventStub = sinon.stub(ExtTelemetry, "sendTelemetryEvent");
    sinon.stub(vscode.commands, "executeCommand");
    const error = new UserError("test source", "test name", "test message", "test displayMessage");
    error.helpLink = "test helpLink";

    await handlers.showError(error);

    chai.assert.isTrue(
      sendTelemetryEventStub.calledWith(extTelemetryEvents.TelemetryEvent.ClickGetHelp, {
        "error-code": "test source.test name",
        "error-message": "test displayMessage",
        "help-link": "test helpLink",
      })
    );
    sinon.restore();
  });

  describe("promptSPFxUpgrade", async () => {
    it("Prompt user to upgrade toolkit when project SPFx version higher than toolkit", async () => {
      sinon.stub(globalVariables, "isSPFxProject").value(true);
      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(""));
      sinon
        .stub(commonTools, "getAppSPFxVersion")
        .resolves(`1.${parseInt(SUPPORTED_SPFX_VERSION.split(".")[1]) + 1}.0`);
      const stubShowMessage = sinon.stub().resolves(ok({}));
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: stubShowMessage,
      });

      await handlers.promptSPFxUpgrade();

      chai.assert(stubShowMessage.calledOnce);
      chai.assert.equal(stubShowMessage.args[0].length, 4);
      sinon.restore();
    });

    it("Prompt user to upgrade project when project SPFx version lower than toolkit", async () => {
      sinon.stub(globalVariables, "isSPFxProject").value(true);
      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(""));
      sinon
        .stub(commonTools, "getAppSPFxVersion")
        .resolves(`1.${parseInt(SUPPORTED_SPFX_VERSION.split(".")[1]) - 1}.0`);

      const stubShowMessage = sinon.stub().resolves(ok({}));
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: stubShowMessage,
      });

      await handlers.promptSPFxUpgrade();

      chai.assert(stubShowMessage.calledOnce);
      chai.assert.equal(stubShowMessage.args[0].length, 4);
      sinon.restore();
    });

    it("Dont show notification when project SPFx version is the same with toolkit", async () => {
      sinon.stub(globalVariables, "isSPFxProject").value(true);
      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file(""));
      sinon.stub(commonTools, "getAppSPFxVersion").resolves(SUPPORTED_SPFX_VERSION);
      const stubShowMessage = sinon.stub();
      sinon.stub(extension, "VS_CODE_UI").value({
        showMessage: stubShowMessage,
      });

      await handlers.promptSPFxUpgrade();

      chai.assert.equal(stubShowMessage.callCount, 0);
      sinon.restore();
    });
  });

  describe("getDotnetPathHandler", async () => {
    afterEach(() => {
      sinon.restore();
    });
    it("dotnet is installed", async () => {
      sinon.stub(DepsManager.prototype, "getStatus").resolves([
        {
          name: ".NET Core SDK",
          type: DepsType.Dotnet,
          isInstalled: true,
          command: "",
          details: {
            isLinuxSupported: false,
            installVersion: "",
            supportedVersions: [],
            binFolders: ["dotnet-bin-folder/dotnet"],
          },
        },
      ]);

      const dotnetPath = await handlers.getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}dotnet-bin-folder${path.delimiter}`);
    });

    it("dotnet is not installed", async () => {
      sinon.stub(DepsManager.prototype, "getStatus").resolves([
        {
          name: ".NET Core SDK",
          type: DepsType.Dotnet,
          isInstalled: false,
          command: "",
          details: {
            isLinuxSupported: false,
            installVersion: "",
            supportedVersions: [],
            binFolders: undefined,
          },
        },
      ]);

      const dotnetPath = await handlers.getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}`);
    });

    it("failed to get dotnet path", async () => {
      sinon.stub(DepsManager.prototype, "getStatus").rejects(new Error("failed to get status"));
      const dotnetPath = await handlers.getDotnetPathHandler();
      chai.assert.equal(dotnetPath, `${path.delimiter}`);
    });
  });

  describe("scaffoldFromDeveloperPortalHandler", async () => {
    beforeEach(() => {
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    });
    afterEach(() => {
      sinon.restore();
    });
    it("missing args", async () => {
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);

      const res = await handlers.scaffoldFromDeveloperPortalHandler();

      chai.assert.equal(res.isOk(), true);
      chai.assert.equal(createProgressBar.notCalled, true);
    });

    it("incorrect number of args", async () => {
      const progressHandler = new ProgressHandler("title", 1);
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);

      const res = await handlers.scaffoldFromDeveloperPortalHandler([]);

      chai.assert.equal(res.isOk(), true);
      chai.assert.equal(createProgressBar.notCalled, true);
    });

    it("general error when signing in M365", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").throws("error1");
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sinon.stub(vscode.window, "showErrorMessage");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);
      chai.assert.isTrue(res.isErr());
      chai.assert.isTrue(createProgressBar.calledOnce);
      chai.assert.isTrue(startProgress.calledOnce);
      chai.assert.isTrue(endProgress.calledOnceWithExactly(false));
      chai.assert.isTrue(showErrorMessage.calledOnce);
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "error1");
      }
    });

    it("error when signing M365", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon
        .stub(M365TokenInstance, "signInWhenInitiatedFromTdp")
        .resolves(err(new UserError("source", "name", "message", "displayMessage")));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sinon.stub(vscode.window, "showErrorMessage");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);

      chai.assert.equal(res.isErr(), true);
      chai.assert.equal(createProgressBar.calledOnce, true);
      chai.assert.equal(startProgress.calledOnce, true);
      chai.assert.equal(endProgress.calledOnceWithExactly(false), true);
      chai.assert.equal(showErrorMessage.calledOnce, true);
    });

    it("error when signing in M365 but missing display message", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon
        .stub(M365TokenInstance, "signInWhenInitiatedFromTdp")
        .resolves(err(new UserError("source", "name", "", "")));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      const showErrorMessage = sinon.stub(vscode.window, "showErrorMessage");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);

      chai.assert.equal(res.isErr(), true);
      chai.assert.equal(createProgressBar.calledOnce, true);
      chai.assert.equal(startProgress.calledOnce, true);
      chai.assert.equal(endProgress.calledOnceWithExactly(false), true);
      chai.assert.equal(showErrorMessage.calledOnce, true);
    });

    it("failed to get teams app", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").resolves(ok("token"));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(commonUtils, "isExistingTabApp").returns(Promise.resolve(false));
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(globalState, "globalStateUpdate");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);
      const getApp = sinon.stub(AppStudioClient, "getApp").throws("error");

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId"]);

      chai.assert.isTrue(res.isErr());
      chai.assert.isTrue(getApp.calledOnce);
      chai.assert.isTrue(createProgressBar.calledOnce);
      chai.assert.isTrue(startProgress.calledOnce);
      chai.assert.isTrue(endProgress.calledOnceWithExactly(true));
    });

    it("happy path", async () => {
      sinon.stub(extension, "VS_CODE_UI").value(new VsCodeUI(<vscode.ExtensionContext>{}));
      const progressHandler = new ProgressHandler("title", 1);
      const startProgress = sinon.stub(progressHandler, "start").resolves();
      const endProgress = sinon.stub(progressHandler, "end").resolves();
      sinon.stub(M365TokenInstance, "signInWhenInitiatedFromTdp").resolves(ok("token"));
      const createProgressBar = sinon
        .stub(extension.VS_CODE_UI, "createProgressBar")
        .returns(progressHandler);
      sinon.stub(handlers, "core").value(new MockCore());
      sinon.stub(commonUtils, "isExistingTabApp").returns(Promise.resolve(false));
      const createProject = sinon.spy(handlers.core, "createProject");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(globalState, "globalStateUpdate");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
      };
      sinon.stub(AppStudioClient, "getApp").resolves(appDefinition);

      const res = await handlers.scaffoldFromDeveloperPortalHandler(["appId", "testuser"]);

      chai.assert.equal(createProject.args[0][0].teamsAppFromTdp.teamsAppId, "mock-id");
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(createProgressBar.calledOnce);
      chai.assert.isTrue(startProgress.calledOnce);
      chai.assert.isTrue(endProgress.calledOnceWithExactly(true));
    });
  });

  describe("publishInDeveloperPortalHandler", async () => {
    afterEach(() => {
      sinon.restore();
    });

    it("publish in developer portal", async () => {
      sinon.stub(handlers, "core").value(new MockCore());
      const publish = sinon.spy(handlers.core, "publishInDeveloperPortal");
      sinon.stub(ExtTelemetry, "sendTelemetryEvent");
      sinon.stub(ExtTelemetry, "sendTelemetryErrorEvent");
      sinon.stub(vscode.commands, "executeCommand");
      sinon.stub(vscodeHelper, "checkerEnabled").returns(false);

      const res = await handlers.publishInDeveloperPortalHandler();
      if (res.isErr()) {
        console.log(res.error);
      }
      chai.assert.isTrue(publish.calledOnce);
    });
  });

  describe("installAppInTeams", () => {
    beforeEach(() => {
      sinon.stub(globalVariables, "workspaceUri").value(vscode.Uri.file("path"));
    });

    afterEach(() => {
      sinon.restore();
    });

    it("v3: happ path", async () => {
      sinon.stub(commonTools, "isV3Enabled").returns(true);
      sinon.stub(debugCommonUtils, "getV3TeamsAppId").returns(Promise.resolve("appId"));
      sinon
        .stub(teamsAppInstallation, "showInstallAppInTeamsMessage")
        .returns(Promise.resolve(true));
      const result = await handlers.installAppInTeams();
      chai.assert.equal(result, undefined);
    });

    it("v3: user cancel", async () => {
      sinon.stub(commonTools, "isV3Enabled").returns(true);
      sinon.stub(debugCommonUtils, "getV3TeamsAppId").returns(Promise.resolve("appId"));
      sinon
        .stub(teamsAppInstallation, "showInstallAppInTeamsMessage")
        .returns(Promise.resolve(false));
      sinon.stub(taskHandler, "terminateAllRunningTeamsfxTasks").callsFake(() => {});
      sinon.stub(debugCommonUtils, "endLocalDebugSession").callsFake(() => {});
      const result = await handlers.installAppInTeams();
      chai.assert.equal(result, "1");
    });

    it("v2: happy path", async () => {
      sinon.stub(commonTools, "isV3Enabled").returns(false);
      sinon.stub(debugCommonUtils, "getDebugConfig").returns(
        Promise.resolve({
          appId: "appId",
          env: "local",
        })
      );
      sinon
        .stub(teamsAppInstallation, "showInstallAppInTeamsMessage")
        .returns(Promise.resolve(true));
      const result = await handlers.installAppInTeams();
      chai.assert.equal(result, undefined);
    });

    it("v2: no appId", async () => {
      sinon.stub(commonTools, "isV3Enabled").returns(false);
      sinon.stub(debugCommonUtils, "getDebugConfig").returns(Promise.resolve(undefined));
      sinon.stub(handlers, "showError").callsFake(async () => {});
      sinon.stub(taskHandler, "terminateAllRunningTeamsfxTasks").callsFake(() => {});
      sinon.stub(debugCommonUtils, "endLocalDebugSession").callsFake(() => {});
      const result = await handlers.installAppInTeams();
      chai.assert.equal(result, "1");
    });
  });
});
