// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { it } from "mocha";
import { TeamsAppSolution } from " ../../../src/plugins/solution";
import {
  ConfigMap,
  FxError,
  ok,
  PluginContext,
  Result,
  SolutionConfig,
  SolutionContext,
  TeamsAppManifest,
  Void,
  Plugin,
  Platform,
} from "@microsoft/teamsfx-api";
import * as sinon from "sinon";
import fs, { PathLike } from "fs-extra";
import {
  BotOptionItem,
  HostTypeOptionSPFx,
  MessageExtensionItem,
  TabOptionItem,
} from "../../../src/plugins/solution/fx-solution/question";
import _ from "lodash";
import path from "path";
import { getTemplatesFolder } from "../../../src";
import { SolutionError } from "../../../src/plugins/solution/fx-solution/constants";
import { validManifest } from "./util";
import * as uuid from "uuid";
import { ResourcePlugins } from "../../../src/plugins/solution/fx-solution/ResourcePluginContainer";
import Container from "typedi";

chai.use(chaiAsPromised);
const expect = chai.expect;
const fehostPlugin = Container.get<Plugin>(ResourcePlugins.FrontendPlugin);
const botPlugin = Container.get<Plugin>(ResourcePlugins.BotPlugin);
function mockSolutionContext(): SolutionContext {
  const config: SolutionConfig = new Map();
  return {
    root: ".",
    // app: new TeamsAppManifest(),
    config,
    answers: { platform: Platform.VSCode },
    projectSettings: undefined,
  };
}

function mockScaffoldThatAlwaysSucceed(plugin: Plugin) {
  plugin.preScaffold = async function (_ctx: PluginContext): Promise<Result<any, FxError>> {
    return ok(Void);
  };
  plugin.scaffold = async function (_ctx: PluginContext): Promise<Result<any, FxError>> {
    return ok(Void);
  };
  plugin.postScaffold = async function (_ctx: PluginContext): Promise<Result<any, FxError>> {
    return ok(Void);
  };
}

describe("Solution scaffold()", () => {
  const mocker = sinon.createSandbox();

  afterEach(() => {
    mocker.restore();
  });

  it("should return error for invalid plugin names", async () => {
    const solution = new TeamsAppSolution();
    const mockedCtx = mockSolutionContext();
    mockedCtx.projectSettings = {
      appName: "my app",
      currentEnv: "default",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionSPFx.id,
        name: "azure",
        version: "1.0",
        activeResourcePlugins: ["SomeInvalidPluginName"],
      },
    };
    const result = await solution.scaffold(mockedCtx);
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().name).equals(SolutionError.PluginNotFound);
  });

  it("should return error if manifest file is not found", async () => {
    const solution = new TeamsAppSolution();
    const mockedCtx = mockSolutionContext();
    mockedCtx.projectSettings = {
      appName: "my app",
      currentEnv: "default",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionSPFx.id,
        name: "azure",
        version: "1.0",
        activeResourcePlugins: [fehostPlugin.name],
      },
    };
    // We leverage the fact that in testing env, this is not file at `${ctx.root}/.${ConfigFolderName}/${REMOTE_MANIFEST}`
    // So we even don't need to mock fs.readJson
    const result = await solution.scaffold(mockedCtx);
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().name).equals("ManifestLoadFailed");
  });
});

describe("Solution scaffold() reading manifest file with no app name", () => {
  const mocker = sinon.createSandbox();
  const fileContent: Map<string, any> = new Map();

  const manifestWithNoAppName = _.cloneDeep(validManifest);
  manifestWithNoAppName.name.short = "";

  beforeEach(() => {
    mocker.stub(fs, "writeFile").callsFake((path: number | PathLike, data: any) => {
      fileContent.set(path.toString(), data);
    });
    // mocker.stub(fs, "writeFile").resolves();
    mocker.stub(fs, "writeJSON").callsFake((file: string, obj: any) => {
      fileContent.set(file, JSON.stringify(obj));
    });
    mocker.stub(fs, "readJson").resolves(manifestWithNoAppName);
    // Uses stub<any, any> to circumvent type check. Beacuse sinon fails to mock my target overload.
    mocker.stub<any, any>(fs, "copy").resolves();
  });

  afterEach(() => {
    mocker.restore();
  });

  it("should return error", async () => {
    const solution = new TeamsAppSolution();
    const mockedCtx = mockSolutionContext();
    mockedCtx.projectSettings = {
      appName: "my app",
      currentEnv: "default",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionSPFx.id,
        name: "azure",
        version: "1.0",
        activeResourcePlugins: [fehostPlugin.name],
      },
    };
    const result = await solution.scaffold(mockedCtx);
    expect(result.isErr()).to.be.true;
    expect(result._unsafeUnwrapErr().name).equals("ManifestLoadFailed");
    expect(result._unsafeUnwrapErr().message).equals(
      "Failed to read manifest file. Error: Name is missing."
    );
  });
});

describe("Solution scaffold() reading valid manifest file", () => {
  const mocker = sinon.createSandbox();
  const fileContent: Map<string, any> = new Map();
  const readmePath = path.join(getTemplatesFolder(), "plugins", "solution", "README.md");
  const mockedReadMeContent = "mocked readme content";

  beforeEach(() => {
    mocker.stub(fs, "writeFile").callsFake((path: number | PathLike, data: any) => {
      fileContent.set(path.toString(), data);
    });
    mocker.stub(fs, "writeJSON").callsFake((file: string, obj: any) => {
      fileContent.set(file, JSON.stringify(obj));
    });
    mocker.stub(fs, "readJson").resolves(validManifest);
    mocker.stub<any, any>(fs, "pathExists").withArgs(readmePath).resolves(true);
    mocker.stub(fs, "copy").callsFake((src: string, dest: string) => {
      fileContent.set(dest, mockedReadMeContent);
    });
  });

  afterEach(() => {
    mocker.restore();
  });

  it("should work for happy path with only tab", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedCtx = mockSolutionContext();
    mockedCtx.projectSettings = {
      appName: "my app",
      currentEnv: "default",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionSPFx.id,
        name: "azure",
        version: "1.0",
        activeResourcePlugins: [fehostPlugin.name],
        capabilities: [TabOptionItem.id],
      },
    };
    mockScaffoldThatAlwaysSucceed(fehostPlugin);

    const result = await solution.scaffold(mockedCtx);
    expect(result.isOk()).to.be.true;
  });

  it("should work and generate README.md for happy path with tab and bot", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedCtx = mockSolutionContext();
    mockedCtx.projectSettings = {
      appName: "my app",
      currentEnv: "default",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionSPFx.id,
        name: "azure",
        version: "1.0",
        activeResourcePlugins: [fehostPlugin.name, botPlugin.name],
        capabilities: [TabOptionItem.id, BotOptionItem.id],
      },
    };
    mockScaffoldThatAlwaysSucceed(fehostPlugin);
    mockScaffoldThatAlwaysSucceed(botPlugin);
    const result = await solution.scaffold(mockedCtx);
    expect(result.isOk()).to.be.true;
    expect(fileContent.get(`${mockedCtx.root}/README.md`)).equals(mockedReadMeContent);
  });

  it("should work and generate README.md for happy path with tab and msgext", async () => {
    fileContent.clear();
    const solution = new TeamsAppSolution();
    const mockedCtx = mockSolutionContext();
    mockedCtx.projectSettings = {
      appName: "my app",
      currentEnv: "default",
      projectId: uuid.v4(),
      solutionSettings: {
        hostType: HostTypeOptionSPFx.id,
        name: "azure",
        version: "1.0",
        activeResourcePlugins: [fehostPlugin.name, botPlugin.name],
        capabilities: [TabOptionItem.id, MessageExtensionItem.id],
      },
    };
    mockScaffoldThatAlwaysSucceed(fehostPlugin);
    mockScaffoldThatAlwaysSucceed(botPlugin);
    const result = await solution.scaffold(mockedCtx);
    expect(result.isOk()).to.be.true;
    expect(fileContent.get(`${mockedCtx.root}/README.md`)).equals(mockedReadMeContent);
  });
});
