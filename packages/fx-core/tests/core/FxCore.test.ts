// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Inputs,
  Platform,
  Stage,
  Ok,
  Err,
  FxError,
  UserError,
  SystemError,
  err,
  ok,
  Result,
  Void,
  LogProvider,
} from "@microsoft/teamsfx-api";
import { assert, expect } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import { FxCore } from "../../src";
import * as featureFlags from "../../src/common/featureFlags";
import { validateProjectSettings } from "../../src/common/projectSettingsHelper";
import { environmentManager } from "../../src/core/environment";
import { setTools } from "../../src/core/globalVars";
import { loadProjectSettings } from "../../src/core/middleware/projectSettingsLoader";
import {
  CoreQuestionNames,
  ProgrammingLanguageQuestion,
  ScratchOptionYesVSC,
} from "../../src/core/question";
import {
  BotOptionItem,
  MessageExtensionItem,
  TabOptionItem,
  TabSPFxItem,
} from "../../src/component/constants";
import { deleteFolder, MockTools, randomAppName } from "./utils";
import * as templateActions from "../../src/common/template-utils/templatesActions";
import { UpdateAadAppDriver } from "../../src/component/driver/aad/update";
import AdmZip from "adm-zip";
import { NoAadManifestExistError } from "../../src/core/error";
import "../../src/component/driver/aad/update";
import { envUtil } from "../../src/component/utils/envUtil";
import { YamlParser } from "../../src/component/configManager/parser";
import {
  DriverDefinition,
  DriverInstance,
  ExecutionError,
  ExecutionOutput,
  ExecutionResult,
  ILifecycle,
  LifecycleName,
  Output,
  UnresolvedPlaceholders,
} from "../../src/component/configManager/interface";
import { DriverContext } from "../../src/component/driver/interface/commonArgs";
import { Readable, Writable } from "stream";
import { coordinator } from "../../src/component/coordinator";

describe("Core basic APIs", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  let appName = randomAppName();
  let projectPath = path.resolve(os.tmpdir(), appName);
  beforeEach(() => {
    setTools(tools);
    sandbox.stub<any, any>(featureFlags, "isPreviewFeaturesEnabled").returns(true);
    sandbox.stub<any, any>(templateActions, "scaffoldFromTemplates").resolves();
  });
  afterEach(async () => {
    sandbox.restore();
    deleteFolder(projectPath);
  });
  describe("create from new", async () => {
    it("CLI with folder input", async () => {
      appName = randomAppName();
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        [CoreQuestionNames.Folder]: os.tmpdir(),
        [CoreQuestionNames.AppName]: appName,
        [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
        [CoreQuestionNames.ProgrammingLanguage]: "javascript",
        [CoreQuestionNames.Capabilities]: ["Tab"],
        stage: Stage.create,
      };
      const res = await core.createProject(inputs);
      projectPath = path.resolve(os.tmpdir(), appName);
      assert.isTrue(res.isOk() && res.value === projectPath);
    });

    it("VSCode without customized default root directory", async () => {
      appName = randomAppName();
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [CoreQuestionNames.AppName]: appName,
        [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
        [CoreQuestionNames.ProgrammingLanguage]: "javascript",
        [CoreQuestionNames.Capabilities]: ["Tab"],
        [CoreQuestionNames.Folder]: os.tmpdir(),
        stage: Stage.create,
      };
      const res = await core.createProject(inputs);
      projectPath = inputs.projectPath!;
      assert.isTrue(res.isOk() && res.value === projectPath);
      const projectSettingsResult = await loadProjectSettings(inputs, true);
      assert.isTrue(projectSettingsResult.isOk());
      if (projectSettingsResult.isOk()) {
        const projectSettings = projectSettingsResult.value;
        const validSettingsResult = validateProjectSettings(projectSettings);
        assert.isTrue(validSettingsResult === undefined);
        assert.isTrue(projectSettings.version === "2.1.0");
      }
    });

    it("VSCode without customized default root directory - new UI", async () => {
      appName = randomAppName();
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [CoreQuestionNames.AppName]: appName,
        [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
        [CoreQuestionNames.ProgrammingLanguage]: "javascript",
        [CoreQuestionNames.Capabilities]: "Tab",
        [CoreQuestionNames.Folder]: os.tmpdir(),
        stage: Stage.create,
      };
      const res = await core.createProject(inputs);
      projectPath = inputs.projectPath!;
      assert.isTrue(res.isOk() && res.value === projectPath);
      const projectSettingsResult = await loadProjectSettings(inputs, true);
      assert.isTrue(projectSettingsResult.isOk());
      if (projectSettingsResult.isOk()) {
        const projectSettings = projectSettingsResult.value;
        const validSettingsResult = validateProjectSettings(projectSettings);
        assert.isTrue(validSettingsResult === undefined);
        assert.isTrue(projectSettings.version === "2.1.0");
      }
    });
  });

  it("scaffold and createEnv, activateEnv", async () => {
    appName = randomAppName();
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      [CoreQuestionNames.AppName]: appName,
      [CoreQuestionNames.Folder]: os.tmpdir(),
      [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
      [CoreQuestionNames.ProgrammingLanguage]: "javascript",
      [CoreQuestionNames.Capabilities]: "Tab",
      stage: Stage.create,
    };
    const createRes = await core.createProject(inputs);
    assert.isTrue(createRes.isOk());
    projectPath = inputs.projectPath!;
    await fs.writeFile(
      path.resolve(projectPath, "templates", "appPackage", "manifest.template.json"),
      "{}"
    );
    const newEnvName = "newEnv";
    const envListResult = await environmentManager.listRemoteEnvConfigs(projectPath);
    if (envListResult.isErr()) {
      assert.fail("failed to list env names");
    }
    assert.isTrue(envListResult.value.length === 1);
    assert.isTrue(envListResult.value[0] === environmentManager.getDefaultEnvName());
    inputs[CoreQuestionNames.NewTargetEnvName] = newEnvName;
    const createEnvRes = await core.createEnv(inputs);
    assert.isTrue(createEnvRes.isOk());

    const newEnvListResult = await environmentManager.listRemoteEnvConfigs(projectPath);
    if (newEnvListResult.isErr()) {
      assert.fail("failed to list env names");
    }
    assert.isTrue(newEnvListResult.value.length === 2);
    assert.isTrue(newEnvListResult.value[0] === environmentManager.getDefaultEnvName());
    assert.isTrue(newEnvListResult.value[1] === newEnvName);

    inputs.env = "newEnv";
    const activateEnvRes = await core.activateEnv(inputs);
    assert.isTrue(activateEnvRes.isOk());
  });

  it("deploy aad manifest happy path with param", async () => {
    const restore = mockedEnv({
      TEAMSFX_V3: "true",
    });
    try {
      const core = new FxCore(tools);
      const appName = mockV3Project();
      // sandbox.stub(UpdateAadAppDriver.prototype, "run").resolves(new Ok(new Map()));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [CoreQuestionNames.AppName]: appName,
        [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
        [CoreQuestionNames.ProgrammingLanguage]: "javascript",
        [CoreQuestionNames.Capabilities]: ["Tab", "TabSSO"],
        [CoreQuestionNames.Folder]: os.tmpdir(),
        stage: Stage.deployAad,
        projectPath: path.join(os.tmpdir(), appName, "samples-v3"),
      };

      const runSpy = sandbox.spy(UpdateAadAppDriver.prototype, "run");
      await core.deployAadManifest(inputs);
      sandbox.assert.calledOnce(runSpy);
      assert.isNotNull(runSpy.getCall(0).args[0]);
      assert.strictEqual(
        runSpy.getCall(0).args[0].manifestTemplatePath,
        path.join(os.tmpdir(), appName, "samples-v3", "aad.manifest.template.json")
      );
      runSpy.restore();
    } finally {
      restore();
    }
  });

  it("deploy aad manifest happy path", async () => {
    const restore = mockedEnv({
      TEAMSFX_V3: "true",
    });
    try {
      const core = new FxCore(tools);
      const appName = mockV3Project();
      sandbox.stub(UpdateAadAppDriver.prototype, "run").resolves(new Ok(new Map()));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [CoreQuestionNames.AppName]: appName,
        [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
        [CoreQuestionNames.ProgrammingLanguage]: "javascript",
        [CoreQuestionNames.Capabilities]: ["Tab", "TabSSO"],
        [CoreQuestionNames.Folder]: os.tmpdir(),
        stage: Stage.deployAad,
        projectPath: path.join(os.tmpdir(), appName, "samples-v3"),
      };
      const res = await core.deployAadManifest(inputs);
      assert.isTrue(await fs.pathExists(path.join(os.tmpdir(), appName, "samples-v3", "build")));
      await deleteTestProject(appName);
      assert.isTrue(res.isOk());
    } finally {
      restore();
    }
  });

  it("deploy aad manifest return err", async () => {
    const restore = mockedEnv({
      TEAMSFX_V3: "true",
    });
    try {
      const core = new FxCore(tools);
      const appName = mockV3Project();
      const appManifestPath = path.join(
        os.tmpdir(),
        appName,
        "samples-v3",
        "aad.manifest.template.json"
      );
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [CoreQuestionNames.AppName]: appName,
        [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
        [CoreQuestionNames.ProgrammingLanguage]: "javascript",
        [CoreQuestionNames.Capabilities]: ["Tab", "TabSSO"],
        [CoreQuestionNames.Folder]: os.tmpdir(),
        stage: Stage.deployAad,
        projectPath: path.join(os.tmpdir(), appName, "samples-v3"),
      };
      sandbox
        .stub(UpdateAadAppDriver.prototype, "run")
        .throws(new UserError("error name", "fake_error", "fake_err_msg"));
      const errMsg = `AAD manifest doesn't exist in ${appManifestPath}, please use the CLI to specify an AAD manifest to deploy.`;
      const res = await core.deployAadManifest(inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.strictEqual(res.error.message, "fake_err_msg");
      }
    } finally {
      restore();
    }
  });

  it("deploy aad manifest not exist", async () => {
    const restore = mockedEnv({
      TEAMSFX_V3: "true",
    });
    try {
      const core = new FxCore(tools);
      const appName = mockV3Project();
      const appManifestPath = path.join(
        os.tmpdir(),
        appName,
        "samples-v3",
        "aad.manifest.template.json"
      );
      await fs.remove(appManifestPath);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [CoreQuestionNames.AppName]: appName,
        [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
        [CoreQuestionNames.ProgrammingLanguage]: "javascript",
        [CoreQuestionNames.Capabilities]: ["Tab", "TabSSO"],
        [CoreQuestionNames.Folder]: os.tmpdir(),
        stage: Stage.deployAad,
        projectPath: path.join(os.tmpdir(), appName, "samples-v3"),
      };
      const errMsg = `AAD manifest doesn't exist in ${appManifestPath}, please use the CLI to specify an AAD manifest to deploy.`;
      const res = await core.deployAadManifest(inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof NoAadManifestExistError);
        assert.equal(res.error.message, errMsg);
      }
      await deleteTestProject(appName);
    } finally {
      restore();
    }
  });

  it("ProgrammingLanguageQuestion", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      [CoreQuestionNames.Capabilities]: TabSPFxItem.id,
    };
    if (
      ProgrammingLanguageQuestion.dynamicOptions &&
      ProgrammingLanguageQuestion.placeholder &&
      typeof ProgrammingLanguageQuestion.placeholder === "function"
    ) {
      const options = ProgrammingLanguageQuestion.dynamicOptions(inputs);
      assert.deepEqual([{ id: "typescript", label: "TypeScript" }], options);
      const placeholder = ProgrammingLanguageQuestion.placeholder(inputs);
      assert.equal("SPFx is currently supporting TypeScript only.", placeholder);
    }

    languageAssert({
      platform: Platform.VSCode,
      [CoreQuestionNames.Capabilities]: TabOptionItem.id,
    });
    languageAssert({
      platform: Platform.VSCode,
      [CoreQuestionNames.Capabilities]: BotOptionItem.id,
    });
    languageAssert({
      platform: Platform.VSCode,
      [CoreQuestionNames.Capabilities]: MessageExtensionItem.id,
    });

    function languageAssert(inputs: Inputs) {
      if (
        ProgrammingLanguageQuestion.dynamicOptions &&
        ProgrammingLanguageQuestion.placeholder &&
        typeof ProgrammingLanguageQuestion.placeholder === "function"
      ) {
        const options = ProgrammingLanguageQuestion.dynamicOptions(inputs);
        assert.deepEqual(
          [
            { id: "javascript", label: "JavaScript" },
            { id: "typescript", label: "TypeScript" },
          ],
          options
        );
        const placeholder = ProgrammingLanguageQuestion.placeholder(inputs);
        assert.equal("Select a programming language.", placeholder);
      }
    }
  });
});

describe("apply yaml template", async () => {
  const tools = new MockTools();
  beforeEach(() => {
    setTools(tools);
  });
  describe("when run with missing input", async () => {
    it("should return error when projectPath is undefined", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: undefined,
      };
      const res = await core.apply(inputs, "", "provision");
      assert.isTrue(
        res.isErr() &&
          res.error.name === "InvalidInput" &&
          res.error.message.includes("projectPath")
      );
    });

    it("should return error when env is undefined", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: undefined,
      };
      const res = await core.apply(inputs, "", "provision");
      assert.isTrue(
        res.isErr() && res.error.name === "InvalidInput" && res.error.message.includes("env")
      );
    });
  });

  describe("when readEnv returns error", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(err(mockedError));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error too", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });

  describe("when YamlParser returns error", async () => {
    const sandbox = sinon.createSandbox();

    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(err(mockedError));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error too", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });

  describe("when running against an empty yaml file", async () => {
    const sandbox = sinon.createSandbox();

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(ok({}));
    });

    after(() => {
      sandbox.restore();
    });

    it("should return ok", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isOk());
    });
  });

  describe("when lifecycle returns error", async () => {
    const sandbox = sinon.createSandbox();
    const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

    class MockedProvision implements ILifecycle {
      name: LifecycleName = "provision";
      driverDefs: DriverDefinition[] = [];
      public async run(ctx: DriverContext): Promise<Result<Output, FxError>> {
        return err(mockedError);
      }

      public resolvePlaceholders(): UnresolvedPlaceholders {
        return [];
      }

      public async execute(ctx: DriverContext): Promise<ExecutionResult> {
        return {
          result: err({
            kind: "Failure",
            error: mockedError,
          }),
          summaries: [],
        };
      }

      public resolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
        return ok([]);
      }
    }

    before(() => {
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(YamlParser.prototype, "parse").resolves(
        ok({
          provision: new MockedProvision(),
        })
      );
    });

    after(() => {
      sandbox.restore();
    });

    it("should return error", async () => {
      const core = new FxCore(tools);
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        env: "dev",
      };
      const res = await core.apply(inputs, "./", "provision");
      assert.isTrue(res.isErr() && res.error.name === "mockedError");
    });
  });
});

function mockV3Project(): string {
  const zip = new AdmZip(path.join(__dirname, "./samples_v3.zip"));
  const appName = randomAppName();
  zip.extractAllTo(path.join(os.tmpdir(), appName));
  return appName;
}

async function deleteTestProject(appName: string) {
  await fs.remove(path.join(os.tmpdir(), appName));
}

describe("createEnvCopyV3", async () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();
  const sourceEnvContent = [
    "# this is a comment",
    "TEAMSFX_ENV=dev",
    "",
    "_KEY1=value1",
    "KEY2=value2",
    "SECRET_KEY3=xxxx",
  ];
  const sourceEnvStr = sourceEnvContent.join(os.EOL);

  const writeStreamContent: string[] = [];
  // fs.WriteStream's full interface is too hard to mock. We only use write() and end() so we just mock them here.
  class MockedWriteStream {
    write(chunk: any, callback?: ((error: Error | null | undefined) => void) | undefined): boolean {
      writeStreamContent.push(chunk);
      return true;
    }
    end(): boolean {
      return true;
    }
  }

  before(() => {
    sandbox.stub(fs, "readFile").resolves(Buffer.from(sourceEnvStr, "utf8"));
    sandbox.stub<any, any>(fs, "createWriteStream").returns(new MockedWriteStream());
  });

  after(() => {
    sandbox.restore();
  });

  it("should create new .env file with desired content", async () => {
    const core = new FxCore(tools);
    const res = await core.createEnvCopyV3("newEnv", "dev", "./");
    assert(res.isOk());
    assert(
      writeStreamContent[0] === `${sourceEnvContent[0]}${os.EOL}`,
      "comments should be copied"
    );
    assert(
      writeStreamContent[1] === `TEAMSFX_ENV=newEnv${os.EOL}`,
      "TEAMSFX_ENV's value should be new env name"
    );
    assert(writeStreamContent[2] === `${os.EOL}`, "empty line should be coped");
    assert(
      writeStreamContent[3] === `_KEY1=${os.EOL}`,
      "key starts with _ should be copied with empty value"
    );
    assert(
      writeStreamContent[4] === `KEY2=${os.EOL}`,
      "key not starts with _ should be copied with empty value"
    );
    assert(
      writeStreamContent[5] === `SECRET_KEY3=${os.EOL}`,
      "key not starts with SECRET_ should be copied with empty value"
    );
  });
});

describe("publishInDeveloperPortal", () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();

  before(() => {
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
  });
  afterEach(() => {
    sandbox.restore();
  });

  it("success", async () => {
    const core = new FxCore(tools);
    const inputs: Inputs = {
      env: "local",
      projectPath: "project-path",
      platform: Platform.VSCode,
      [CoreQuestionNames.ManifestPath]: "manifest-path",
    };
    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(coordinator, "publishInDeveloperPortal").resolves(ok(Void));
    const res = await core.publishInDeveloperPortal(inputs);

    if (res.isErr()) {
      console.log(res.error);
    }
    assert.isTrue(res.isOk());
  });
});
