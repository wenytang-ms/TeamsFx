// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import * as path from "path";
import { Argv, Options } from "yargs";

import { err, FxError, ok, Result, LogLevel } from "@microsoft/teamsfx-api";

import activate from "../activate";
import { getSystemInputs, readEnvJsonFile, setSubscriptionId } from "../utils";
import { YargsCommand } from "../yargsCommand";
import CliTelemetry from "../telemetry/cliTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import CLIUIInstance from "../userInteraction";
import CLILogProvider from "../commonlib/log";
import HelpParamGenerator from "../helpParamGenerator";

export class ResourceAdd extends YargsCommand {
  public readonly commandHead = `add`;
  public readonly command = `${this.commandHead} <resource-type>`;
  public readonly description = "Add a resource to the current application.";

  public readonly subCommands: YargsCommand[] = [
    new ResourceAddSql(),
    new ResourceAddApim(),
    new ResourceAddFunction(),
  ];

  public builder(yargs: Argv): Argv<any> {
    this.subCommands.forEach((cmd) => {
      yargs.command(cmd.command, cmd.description, cmd.builder.bind(cmd), cmd.handler.bind(cmd));
    });

    return yargs;
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    return ok(null);
  }
}

export class ResourceAddSql extends YargsCommand {
  public readonly commandHead = `azure-sql`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Add a new SQL database.";
  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp("addResource-sql");
    return yargs.options(this.params);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve(args.folder || "./");
    CliTelemetry.withRootFolder(rootFolder).sendTelemetryEvent(TelemetryEvent.UpdateProjectStart, {
      [TelemetryProperty.Resources]: this.commandHead,
    });

    CLIUIInstance.updatePresetAnswers(this.params, args);

    const result = await activate(rootFolder);
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateProject, result.error, {
        [TelemetryProperty.Resources]: this.commandHead,
      });
      return err(result.error);
    }

    const func = {
      namespace: "fx-solution-azure",
      method: "addResource",
    };

    const core = result.value;

    {
      const result = await core.executeUserTask(func, getSystemInputs(rootFolder));
      if (result.isErr()) {
        CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateProject, result.error, {
          [TelemetryProperty.Resources]: this.commandHead,
        });
        return err(result.error);
      }
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.UpdateProject, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.Resources]: this.commandHead,
    });
    return ok(null);
  }
}

export class ResourceAddApim extends YargsCommand {
  public readonly commandHead = `azure-apim`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Add a new API Managment service instance.";
  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp("addResource-apim");
    return yargs.options(this.params);
  }

  public async runCommand(args: {
    [argName: string]: string | undefined;
  }): Promise<Result<null, FxError>> {
    if (!("apim-resource-group" in args)) {
      args["apim-resource-group"] = undefined;
    }
    if (!("apim-service-name" in args)) {
      args["apim-service-name"] = undefined;
    }
    const rootFolder = path.resolve(args.folder || "./");
    CliTelemetry.withRootFolder(rootFolder).sendTelemetryEvent(TelemetryEvent.UpdateProjectStart, {
      [TelemetryProperty.Resources]: this.commandHead,
    });

    CLIUIInstance.updatePresetAnswers(this.params, args);

    {
      const result = await setSubscriptionId(args.subscription, rootFolder);
      if (result.isErr()) {
        CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateProject, result.error, {
          [TelemetryProperty.Resources]: this.commandHead,
        });
        return result;
      }
    }

    const result = await activate(rootFolder);
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateProject, result.error, {
        [TelemetryProperty.Resources]: this.commandHead,
      });
      return err(result.error);
    }

    const func = {
      namespace: "fx-solution-azure",
      method: "addResource",
    };

    const core = result.value;
    {
      const result = await core.executeUserTask(func, getSystemInputs(rootFolder));
      if (result.isErr()) {
        CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateProject, result.error, {
          [TelemetryProperty.Resources]: this.commandHead,
        });
        return err(result.error);
      }
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.UpdateProject, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.Resources]: this.commandHead,
    });
    return ok(null);
  }
}

export class ResourceAddFunction extends YargsCommand {
  public readonly commandHead = `azure-function`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Add a new function app.";
  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp("addResource-function");
    return yargs.options(this.params);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve(args.folder || "./");
    CliTelemetry.withRootFolder(rootFolder).sendTelemetryEvent(TelemetryEvent.UpdateProjectStart, {
      [TelemetryProperty.Resources]: this.commandHead,
    });

    CLIUIInstance.updatePresetAnswers(this.params, args);

    const result = await activate(rootFolder);
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateProject, result.error, {
        [TelemetryProperty.Resources]: this.commandHead,
      });
      return err(result.error);
    }

    const func = {
      namespace: "fx-solution-azure",
      method: "addResource",
    };

    const core = result.value;
    {
      const result = await core.executeUserTask(func, getSystemInputs(rootFolder));
      if (result.isErr()) {
        CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateProject, result.error, {
          [TelemetryProperty.Resources]: this.commandHead,
        });
        return err(result.error);
      }
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.UpdateProject, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.Resources]: this.commandHead,
    });
    return ok(null);
  }
}

export class ResourceShow extends YargsCommand {
  public readonly commandHead = `show`;
  public readonly command = `${this.commandHead} <resource-type>`;
  public readonly description =
    "Show configuration details of resources in the current application.";

  public readonly subCommands: YargsCommand[] = [
    new ResourceShowFunction(),
    new ResourceShowSQL(),
    new ResourceShowApim(),
  ];

  public builder(yargs: Argv): Argv<any> {
    this.subCommands.forEach((cmd) => {
      yargs.command(cmd.command, cmd.description, cmd.builder.bind(cmd), cmd.handler.bind(cmd));
    });

    return yargs;
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    return ok(null);
  }
}

export class ResourceShowFunction extends YargsCommand {
  public readonly commandHead = `azure-function`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Azure Functions details";
  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp("ResourceShowFunction");
    return yargs.options(this.params);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve(args["folder"] || "./");
    const result = await readEnvJsonFile(rootFolder);
    // TODO: should be generated by 'paramGenerator.ts'
    const pluginName = "fx-resource-function";
    if (result.isOk()) {
      const answer: { [_: string]: string } = {};
      if (pluginName in result.value) {
        answer[pluginName] = result.value[pluginName];
      }
      CLILogProvider.necessaryLog(LogLevel.Info, JSON.stringify(answer, undefined, 4), true);
      return ok(null);
    } else {
      return err(result.error);
    }
  }
}

export class ResourceShowSQL extends YargsCommand {
  public readonly commandHead = `azure-sql`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Azure SQL details";
  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp("ResourceShowSQL");
    return yargs.options(this.params);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve(args["folder"] || "./");
    const result = await readEnvJsonFile(rootFolder);
    // TODO: should be generated by 'paramGenerator.ts'
    const pluginName = "fx-resource-azure-sql";
    if (result.isOk()) {
      const answer: { [_: string]: string } = {};
      if (pluginName in result.value) {
        answer[pluginName] = result.value[pluginName];
      }
      CLILogProvider.necessaryLog(LogLevel.Info, JSON.stringify(answer, undefined, 4), true);
      return ok(null);
    } else {
      return err(result.error);
    }
  }
}

export class ResourceShowApim extends YargsCommand {
  public readonly commandHead = `azure-apim`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Azure APIM details";
  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp("ResourceShowApim");
    return yargs.options(this.params);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve(args["folder"] || "./");
    const result = await readEnvJsonFile(rootFolder);
    // TODO: should be generated by 'paramGenerator.ts'
    const pluginName = "fx-resource-apim";
    if (result.isOk()) {
      const answer: { [_: string]: string } = {};
      if (pluginName in result.value) {
        answer[pluginName] = result.value[pluginName];
      }
      CLILogProvider.necessaryLog(LogLevel.Info, JSON.stringify(answer, undefined, 4), true);
      return ok(null);
    } else {
      return err(result.error);
    }
  }
}

export class ResourceList extends YargsCommand {
  public readonly commandHead = `list`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "List all of the resources in the current application";
  public params: { [_: string]: Options } = {};

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp("ResourceList");
    return yargs.options(this.params);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    const rootFolder = path.resolve(args["folder"] || "./");
    const result = await readEnvJsonFile(rootFolder);
    const pluginNameMap: Map<string, string> = new Map();
    pluginNameMap.set("fx-resource-azure-sql", "azure-sql");
    pluginNameMap.set("fx-resource-function", "azure-function");
    pluginNameMap.set("fx-resource-apim", "azure-apim");

    if (result.isOk()) {
      const answer: { [_: string]: string } = {};
      for (const [pluginAlias, _] of pluginNameMap) {
        if (pluginAlias in result.value) {
          answer[pluginAlias] = result.value[pluginAlias];
        }
      }
      CLILogProvider.necessaryLog(LogLevel.Info, JSON.stringify(answer, undefined, 4), true);
      return ok(null);
    } else {
      return err(result.error);
    }
  }
}

export default class Resource extends YargsCommand {
  public readonly commandHead = `resource`;
  public readonly command = `${this.commandHead} <action>`;
  public readonly description = "Manage the resources in the current application.";

  public readonly subCommands: YargsCommand[] = [
    new ResourceAdd(),
    new ResourceShow(),
    new ResourceList(),
  ];

  public builder(yargs: Argv): Argv<any> {
    this.subCommands.forEach((cmd) => {
      yargs.command(cmd.command, cmd.description, cmd.builder.bind(cmd), cmd.handler.bind(cmd));
    });
    return yargs.version(false);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    return ok(null);
  }
}
