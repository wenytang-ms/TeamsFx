// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import {
  err,
  FxError,
  Result,
  ok,
  Inputs,
  Platform,
  Stage,
  SolutionContext,
  QTreeNode,
  Func,
  InputTextConfig,
  InputTextResult,
  SelectFolderConfig,
  SelectFolderResult,
  SingleSelectConfig,
  SingleSelectResult,
  OptionItem,
  traverse,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import * as os from "os";
import { FunctionRouterError, FxCore, InvalidInputError, validateProject } from "../../src";
import sinon from "sinon";
import { MockSolution, MockTools, randomAppName } from "./utils";
import { loadSolutionContext } from "../../src/core/middleware/contextLoader";
import { defaultSolutionLoader } from "../../src/core/loader";
import {
  CoreQuestionNames,
  SampleSelect,
  ScratchOptionNoVSC,
  ScratchOptionYesVSC,
} from "../../src/core/question";

describe("Core basic APIs", () => {
  const sandbox = sinon.createSandbox();
  const mockSolution = new MockSolution();
  const tools = new MockTools();
  const ui = tools.ui;
  let appName = randomAppName();
  let projectPath = path.resolve(os.tmpdir(), appName);

  beforeEach(() => {
    sandbox.stub<any, any>(defaultSolutionLoader, "loadSolution").resolves(mockSolution);
    sandbox.stub<any, any>(defaultSolutionLoader, "loadGlobalSolutions").resolves([mockSolution]);
  });

  afterEach(async () => {
    sandbox.restore();
    await fs.rmdir(projectPath, { recursive: true });
  });

  it("happy path: create from new, provision, deploy, localDebug, publish, getQuestion, getQuestionsForUserTask, getProjectConfig, setSubscriptionInfo", async () => {
    const expectedInputs: Inputs = {
      platform: Platform.CLI,
      [CoreQuestionNames.AppName]: appName,
      [CoreQuestionNames.Foler]: os.tmpdir(),
      [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
      projectPath: projectPath,
      solution: mockSolution.name,
    };
    sandbox
      .stub<any, any>(ui, "inputText")
      .callsFake(async (config: InputTextConfig): Promise<Result<InputTextResult, FxError>> => {
        if (config.name === CoreQuestionNames.AppName) {
          return ok({
            type: "success",
            result: expectedInputs[CoreQuestionNames.AppName] as string,
          });
        }
        throw err(InvalidInputError("invalid question"));
      });
    sandbox
      .stub<any, any>(ui, "selectFolder")
      .callsFake(
        async (config: SelectFolderConfig): Promise<Result<SelectFolderResult, FxError>> => {
          if (config.name === CoreQuestionNames.Foler) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.Foler] as string,
            });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );
    sandbox
      .stub<any, any>(ui, "selectOption")
      .callsFake(
        async (config: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> => {
          if (config.name === CoreQuestionNames.CreateFromScratch) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.CreateFromScratch] as string,
            });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );
    const core = new FxCore(tools);
    {
      const inputs: Inputs = { platform: Platform.CLI };
      const res = await core.createProject(inputs);
      assert.isTrue(res.isOk() && res.value === projectPath);
      assert.deepEqual(expectedInputs, inputs);
      const loadRes = await loadSolutionContext(tools, inputs);
      assert.isTrue(loadRes.isOk());
      if (loadRes.isOk()) {
        const solutionContext = loadRes.value;
        const validRes = validateProject(solutionContext);
        assert.isTrue(validRes === undefined);
        const solutioConfig = solutionContext.config.get("solution");
        assert.isTrue(solutioConfig !== undefined);
        assert.isTrue(solutioConfig!.get("create") === true);
        assert.isTrue(solutioConfig!.get("scaffold") === true);
      }
    }
    {
      const inputs: Inputs = { platform: Platform.CLI, projectPath: projectPath };
      let res = await core.provisionResources(inputs);
      assert.isTrue(res.isOk());

      res = await core.deployArtifacts(inputs);
      assert.isTrue(res.isOk());

      res = await core.localDebug(inputs);
      assert.isTrue(res.isOk());

      res = await core.publishApplication(inputs);
      assert.isTrue(res.isOk());

      const func: Func = { method: "test", namespace: "fx-solution-mock" };
      const res2 = await core.executeUserTask(func, inputs);
      assert.isTrue(res2.isOk());

      const loadRes = await loadSolutionContext(tools, inputs);
      assert.isTrue(loadRes.isOk());
      if (loadRes.isOk()) {
        const solutionContext = loadRes.value;
        const validRes = validateProject(solutionContext);
        assert.isTrue(validRes === undefined);
        const solutioConfig = solutionContext.config.get("solution");
        assert.isTrue(solutioConfig !== undefined);
        assert.isTrue(solutioConfig!.get("provision") === true);
        assert.isTrue(solutioConfig!.get("deploy") === true);
        assert.isTrue(solutioConfig!.get("localDebug") === true);
        assert.isTrue(solutioConfig!.get("publish") === true);
        assert.isTrue(solutioConfig!.get("executeUserTask") === true);
      }
    }

    //getQuestion
    {
      const inputs: Inputs = { platform: Platform.VSCode, projectPath: projectPath };
      const res = await core.getQuestions(Stage.provision, inputs);
      assert.isTrue(res.isOk() && res.value === undefined);
    }
    //getQuestionsForUserTask
    {
      const inputs: Inputs = { platform: Platform.VSCode, projectPath: projectPath };
      const func: Func = { namespace: "fx-solution-mock", method: "mock" };
      const res = await core.getQuestionsForUserTask(func, inputs);
      assert.isTrue(res.isOk() && res.value === undefined);
    }
    //getProjectConfig
    {
      const inputs: Inputs = { platform: Platform.VSCode, projectPath: projectPath };
      const res = await core.getProjectConfig(inputs);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        const projectConfig = res.value;
        assert.isTrue(projectConfig !== undefined);
        if (projectConfig !== undefined) {
          assert.isTrue(projectConfig.settings !== undefined);
          assert.isTrue(projectConfig.config !== undefined);
        }
      }
    }
    //setSubscriptionInfo
    {
      const inputs: Inputs = { platform: Platform.VSCode, projectPath: projectPath };
      inputs["subscriptionId"] = "000000-11111";
      inputs["tenantId"] = "222222-33333";
      const res = await core.setSubscriptionInfo(inputs);
      assert.isTrue(res.isOk());

      const inputs2: Inputs = { platform: Platform.VSCode, projectPath: projectPath };
      const res2 = await core.getProjectConfig(inputs2);
      assert.isTrue(res2.isOk());
      if (res2.isOk()) {
        const projectConfig = res2.value;
        assert.isTrue(projectConfig !== undefined);
        if (projectConfig !== undefined) {
          assert.isTrue(projectConfig.settings !== undefined);
          assert.isTrue(projectConfig.config !== undefined);
          const sconfig = projectConfig.config!.get("solution");
          assert.isTrue(
            sconfig !== undefined &&
              sconfig.get("subscriptionId") === "000000-11111" &&
              sconfig.get("tenantId") === "222222-33333"
          );
        }
      }
    }

    {
      const inputs: Inputs = { platform: Platform.VSCode, projectPath: projectPath };
      const res = await core.setSubscriptionInfo(inputs);
      assert.isTrue(res.isOk());

      const inputs2: Inputs = { platform: Platform.VSCode, projectPath: projectPath };
      const res2 = await core.getProjectConfig(inputs2);
      assert.isTrue(res2.isOk());
      if (res2.isOk()) {
        const projectConfig = res2.value;
        assert.isTrue(projectConfig !== undefined);
        if (projectConfig !== undefined) {
          assert.isTrue(projectConfig.settings !== undefined);
          assert.isTrue(projectConfig.config !== undefined);
          const sconfig = projectConfig.config!.get("solution");
          assert.isTrue(
            sconfig !== undefined &&
              sconfig.get("subscriptionId") === undefined &&
              sconfig.get("tenantId") === undefined
          );
        }
      }
    }
  });

  it("happy path: create from sample", async () => {
    const sampleOption = SampleSelect.staticOptions[0] as OptionItem;
    appName = sampleOption.id;
    projectPath = path.resolve(os.tmpdir(), appName);
    const expectedInputs: Inputs = {
      platform: Platform.CLI,
      [CoreQuestionNames.Foler]: os.tmpdir(),
      [CoreQuestionNames.CreateFromScratch]: ScratchOptionNoVSC.id,
      [CoreQuestionNames.Samples]: sampleOption,
    };
    sandbox
      .stub<any, any>(ui, "selectFolder")
      .callsFake(
        async (config: SelectFolderConfig): Promise<Result<SelectFolderResult, FxError>> => {
          if (config.name === CoreQuestionNames.Foler) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.Foler] as string,
            });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );
    sandbox
      .stub<any, any>(ui, "selectOption")
      .callsFake(
        async (config: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> => {
          if (config.name === CoreQuestionNames.CreateFromScratch) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.CreateFromScratch] as string,
            });
          }
          if (config.name === CoreQuestionNames.Samples) {
            return ok({ type: "success", result: sampleOption });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );
    const core = new FxCore(tools);
    {
      const inputs: Inputs = { platform: Platform.CLI };
      const res = await core.createProject(inputs);
      assert.isTrue(res.isOk() && res.value === projectPath);
      assert.deepEqual(expectedInputs, inputs);
      inputs.projectPath = projectPath;
      const loadRes = await loadSolutionContext(tools, inputs);
      assert.isTrue(loadRes.isOk());
      if (loadRes.isOk()) {
        const solutionContext = loadRes.value;
        const validRes = validateProject(solutionContext);
        assert.isTrue(validRes === undefined);
        const solutioConfig = solutionContext.config.get("solution");
        assert.isTrue(solutioConfig !== undefined);
      }
    }
    {
      const inputs: Inputs = { platform: Platform.CLI, projectPath: projectPath };
      let res = await core.provisionResources(inputs);
      assert.isTrue(res.isOk());

      res = await core.deployArtifacts(inputs);
      assert.isTrue(res.isOk());

      res = await core.localDebug(inputs);
      assert.isTrue(res.isOk());

      res = await core.publishApplication(inputs);
      assert.isTrue(res.isOk());

      const func: Func = { method: "test", namespace: "fx-solution-mock" };
      const res2 = await core.executeUserTask(func, inputs);
      assert.isTrue(res2.isOk());
      const loadRes = await loadSolutionContext(tools, inputs);
      assert.isTrue(loadRes.isOk());
      if (loadRes.isOk()) {
        const solutionContext = loadRes.value;
        const validRes = validateProject(solutionContext);
        assert.isTrue(validRes === undefined);
        const solutioConfig = solutionContext.config.get("solution");
        assert.isTrue(solutioConfig !== undefined);
        assert.isTrue(solutioConfig!.get("provision") === true);
        assert.isTrue(solutioConfig!.get("deploy") === true);
        assert.isTrue(solutioConfig!.get("localDebug") === true);
        assert.isTrue(solutioConfig!.get("publish") === true);
        assert.isTrue(solutioConfig!.get("executeUserTask") === true);
      }
    }
  });

  it("happy path: getQuestions for create", async () => {
    const expectedInputs: Inputs = {
      platform: Platform.CLI,
      [CoreQuestionNames.AppName]: appName,
      [CoreQuestionNames.Foler]: os.tmpdir(),
      [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
      solution: mockSolution.name,
    };
    sandbox
      .stub<any, any>(ui, "inputText")
      .callsFake(async (config: InputTextConfig): Promise<Result<InputTextResult, FxError>> => {
        if (config.name === CoreQuestionNames.AppName) {
          return ok({
            type: "success",
            result: expectedInputs[CoreQuestionNames.AppName] as string,
          });
        }
        throw err(InvalidInputError("invalid question"));
      });
    sandbox
      .stub<any, any>(ui, "selectFolder")
      .callsFake(
        async (config: SelectFolderConfig): Promise<Result<SelectFolderResult, FxError>> => {
          if (config.name === CoreQuestionNames.Foler) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.Foler] as string,
            });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );
    sandbox
      .stub<any, any>(ui, "selectOption")
      .callsFake(
        async (config: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> => {
          if (config.name === CoreQuestionNames.CreateFromScratch) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.CreateFromScratch] as string,
            });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );

    const core = new FxCore(tools);
    const inputs: Inputs = { platform: Platform.CLI };
    const res = await core.getQuestions(Stage.create, inputs);
    assert.isTrue(res.isOk());

    if (res.isOk()) {
      const node = res.value;
      if (node) {
        const traverseRes = await traverse(node, inputs, ui);
        assert.isTrue(traverseRes.isOk());
      }
      assert.deepEqual(expectedInputs, inputs);
    }
  });

  it("happy path: getQuestions, getQuestionsForUserTask for static question", async () => {
    const core = new FxCore(tools);
    {
      const inputs: Inputs = { platform: Platform.VS };
      const res = await core.getQuestions(Stage.provision, inputs);
      assert.isTrue(res.isOk() && res.value === undefined);
    }
    {
      const inputs: Inputs = { platform: Platform.CLI_HELP };
      const res = await core.getQuestions(Stage.provision, inputs);
      assert.isTrue(res.isOk() && res.value === undefined);
    }
    {
      const inputs: Inputs = { platform: Platform.VS };
      const func: Func = { namespace: "fx-solution-mock", method: "mock" };
      const res = await core.getQuestionsForUserTask(func, inputs);
      assert.isTrue(res.isOk() && res.value === undefined);
    }
    {
      const inputs: Inputs = { platform: Platform.CLI_HELP };
      const func: Func = { namespace: "fx-solution-mock", method: "mock" };
      const res = await core.getQuestionsForUserTask(func, inputs);
      assert.isTrue(res.isOk() && res.value === undefined);
    }
    {
      const inputs: Inputs = { platform: Platform.CLI_HELP };
      const func: Func = { namespace: "", method: "mock" };
      const res = await core.getQuestionsForUserTask(func, inputs);
      assert.isTrue(res.isErr() && res.error.name === FunctionRouterError(func).name);
    }

    sandbox
      .stub<any, any>(mockSolution, "getQuestions")
      .callsFake(
        async (
          task: Stage,
          ctx: SolutionContext
        ): Promise<Result<QTreeNode | undefined, FxError>> => {
          return ok(
            new QTreeNode({
              type: "text",
              name: "mock-question",
              title: "mock-question",
            })
          );
        }
      );
    sandbox
      .stub<any, any>(mockSolution, "getQuestionsForUserTask")
      .callsFake(
        async (
          func: Func,
          ctx: SolutionContext
        ): Promise<Result<QTreeNode | undefined, FxError>> => {
          return ok(
            new QTreeNode({
              type: "text",
              name: "mock-question-user-task",
              title: "mock-question-user-task",
            })
          );
        }
      );

    {
      const inputs: Inputs = { platform: Platform.VS };
      const res = await core.getQuestions(Stage.provision, inputs);
      assert.isTrue(res.isOk() && res.value && res.value.data.name === "mock-question");
    }
    {
      const inputs: Inputs = { platform: Platform.CLI_HELP };
      const res = await core.getQuestions(Stage.provision, inputs);
      assert.isTrue(res.isOk() && res.value && res.value.data.name === "mock-question");
    }
    {
      const inputs: Inputs = { platform: Platform.VS };
      const func: Func = { namespace: "fx-solution-mock", method: "mock" };
      const res = await core.getQuestionsForUserTask(func, inputs);
      assert.isTrue(res.isOk() && res.value && res.value.data.name === "mock-question-user-task");
    }
    {
      const inputs: Inputs = { platform: Platform.CLI_HELP };
      const func: Func = { namespace: "fx-solution-mock", method: "mock" };
      const res = await core.getQuestionsForUserTask(func, inputs);
      assert.isTrue(res.isOk() && res.value && res.value.data.name === "mock-question-user-task");
    }
  });

  it("crypto: encrypt, decrypt secrets", async () => {
    appName = randomAppName();
    projectPath = path.resolve(os.tmpdir(), appName);
    const expectedInputs: Inputs = {
      platform: Platform.CLI,
      [CoreQuestionNames.AppName]: appName,
      [CoreQuestionNames.Foler]: os.tmpdir(),
      [CoreQuestionNames.CreateFromScratch]: ScratchOptionYesVSC.id,
      projectPath: projectPath,
      solution: mockSolution.name,
    };
    sandbox
      .stub<any, any>(ui, "inputText")
      .callsFake(async (config: InputTextConfig): Promise<Result<InputTextResult, FxError>> => {
        if (config.name === CoreQuestionNames.AppName) {
          return ok({
            type: "success",
            result: expectedInputs[CoreQuestionNames.AppName] as string,
          });
        }
        throw err(InvalidInputError("invalid question"));
      });
    sandbox
      .stub<any, any>(ui, "selectFolder")
      .callsFake(
        async (config: SelectFolderConfig): Promise<Result<SelectFolderResult, FxError>> => {
          if (config.name === CoreQuestionNames.Foler) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.Foler] as string,
            });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );
    sandbox
      .stub<any, any>(ui, "selectOption")
      .callsFake(
        async (config: SingleSelectConfig): Promise<Result<SingleSelectResult, FxError>> => {
          if (config.name === CoreQuestionNames.CreateFromScratch) {
            return ok({
              type: "success",
              result: expectedInputs[CoreQuestionNames.CreateFromScratch] as string,
            });
          }
          throw err(InvalidInputError("invalid question"));
        }
      );
    const core = new FxCore(tools);
    {
      const inputs: Inputs = { platform: Platform.CLI };
      const res = await core.createProject(inputs);
      assert.isTrue(res.isOk() && res.value === projectPath);

      const encrypted = await core.encrypt("test secret data", inputs);
      assert.isTrue(encrypted.isOk());
      if (encrypted.isOk()) {
        assert.isTrue(encrypted.value.startsWith("crypto_"));
        const decrypted = await core.decrypt(encrypted.value, inputs);
        assert(decrypted.isOk());
        if (decrypted.isOk()) {
          assert.strictEqual(decrypted.value, "test secret data");
        }
      }
    }
  });
});
