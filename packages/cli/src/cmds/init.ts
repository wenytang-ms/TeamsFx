// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { Argv, Options } from "yargs";

import { FxError, err, ok, Result, Func, Platform, LogLevel } from "@microsoft/teamsfx-api";

import { YargsCommand } from "../yargsCommand";
import activate from "../activate";
import CliTelemetry from "../telemetry/cliTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import { argsToInputs } from "../utils";
import CLILogProvider from "../commonlib/log";

export default class Init extends YargsCommand {
  public readonly commandHead = `init`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Add Teams support to an existing Blazor application.";

  public readonly params: { [_: string]: Options } = {
    "app-name": {
      type: "string",
      description: "Application name.",
      default: "TeamsBlazorApp",
    },
    environment: {
      type: "string",
      description: "Environment: 'local' or 'remote'.",
      choices: ["local", "remote"],
      default: "local",
    },
    endpoint: {
      type: "string",
      description: "Teams app endpoint.",
      default: "https://localhost:44357",
    },
    "root-path": {
      type: "string",
      description: "Path to the setting files.",
      default: "./",
    },
  };

  public builder(yargs: Argv): Argv<any> {
    return yargs.version(false).options(this.params);
  }

  public async runCommand(args: {
    [argName: string]: string | string[];
  }): Promise<Result<null, FxError>> {
    const answers = argsToInputs(this.params, args);
    CliTelemetry.sendTelemetryEvent(TelemetryEvent.InitStart);
    const result = await activate();
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Init, result.error);
      return err(result.error);
    }

    const core = result.value;
    {
      answers.platform = Platform.VS;

      const func: Func = {
        namespace: "fx-solution-azure",
        method: "registerTeamsAppAndAad",
      };

      const result = await core.executeUserTask!(func, answers);
      if (result.isErr()) {
        CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Init, result.error);
        return err(result.error);
      }
      CLILogProvider.necessaryLog(LogLevel.Info, JSON.stringify(result.value, null, 4), true);
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.Init, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });
    return ok(null);
  }
}
