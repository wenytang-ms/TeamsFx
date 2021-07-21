// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { Argv, Options } from "yargs";
import { FxError, err, ok, Result, Platform, Func, Stage } from "@microsoft/teamsfx-api";
import activate from "../activate";
import { YargsCommand } from "../yargsCommand";
import { argsToInputs } from "../utils";
import CliTelemetry from "../telemetry/cliTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import HelpParamGenerator from "../helpParamGenerator";

export default class Publish extends YargsCommand {
  public readonly commandHead = `publish`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Publish the app to Teams.";

  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp(Stage.publish);
    return yargs.version(false).options(this.params);
  }

  public async runCommand(args: {
    [argName: string]: string | string[];
  }): Promise<Result<null, FxError>> {
    const answers = argsToInputs(this.params, args);

    const manifestFolderParamName = "manifest-folder";
    let result;
    // if input manifestFolderParam(actually also teams-app-id param),
    // this call is from VS platform, since CLI hide these two param from users.
    if (answers[manifestFolderParamName] && answers["teams-app-id"]) {
      CliTelemetry.sendTelemetryEvent(TelemetryEvent.PublishStart);
      result = await activate();
    } else {
      const rootFolder = answers["folder"] as string;
      delete answers.folder;
      CliTelemetry.withRootFolder(rootFolder).sendTelemetryEvent(TelemetryEvent.PublishStart);
      result = await activate(rootFolder);
    }

    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Publish, result.error);
      return err(result.error);
    }

    const core = result.value;
    if (answers[manifestFolderParamName] && answers["teams-app-id"]) {
      answers.platform = Platform.VS;
      const func: Func = {
        namespace: "fx-solution-azure",
        method: "VSpublish",
      };
      result = await core.executeUserTask!(func, answers);
    } else {
      result = await core.publishApplication(answers);
    }
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Publish, result.error);
      return err(result.error);
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.Publish, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });
    return ok(null);
  }
}
