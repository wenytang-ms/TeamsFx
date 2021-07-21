// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AzureSolutionSettings,
  err,
  Func,
  FxError,
  Plugin,
  PluginContext,
  QTreeNode,
  Result,
  Stage,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import {
  AzureResourceSQL,
  HostTypeOptionAzure,
  TabOptionItem,
} from "../../solution/fx-solution/question";
import { ResourcePlugins } from "../../solution/fx-solution/ResourcePluginContainer";
import { Telemetry } from "./constants";
import { ErrorMessage } from "./errors";
import { SqlPluginImpl } from "./plugin";
import { SqlResult, SqlResultFactory } from "./results";
import { DialogUtils } from "./utils/dialogUtils";
import { TelemetryUtils } from "./utils/telemetryUtils";
@Service(ResourcePlugins.SqlPlugin)
export class SqlPlugin implements Plugin {
  name = "fx-resource-azure-sql";
  displayName = "Azure SQL Datebase";
  activate(solutionSettings: AzureSolutionSettings): boolean {
    const azureResources = solutionSettings.azureResources || [];
    const cap = solutionSettings.capabilities || [];
    return (
      solutionSettings.hostType === HostTypeOptionAzure.id &&
      cap.includes(TabOptionItem.id) &&
      azureResources.includes(AzureResourceSQL.id)
    );
  }
  sqlImpl = new SqlPluginImpl();

  public async preProvision(ctx: PluginContext): Promise<SqlResult> {
    return this.runWithSqlError(
      Telemetry.stage.preProvision,
      () => this.sqlImpl.preProvision(ctx),
      ctx
    );
  }

  public async provision(ctx: PluginContext): Promise<SqlResult> {
    return this.runWithSqlError(Telemetry.stage.provision, () => this.sqlImpl.provision(ctx), ctx);
  }

  public async postProvision(ctx: PluginContext): Promise<SqlResult> {
    return this.runWithSqlError(
      Telemetry.stage.postProvision,
      () => this.sqlImpl.postProvision(ctx),
      ctx
    );
  }

  public async getQuestions(
    stage: Stage,
    ctx: PluginContext
  ): Promise<Result<QTreeNode | undefined, FxError>> {
    return this.runWithSqlError(
      Telemetry.stage.getQuestion,
      () => this.sqlImpl.getQuestions(stage, ctx),
      ctx
    );
  }

  private async runWithSqlError(
    stage: string,
    fn: () => Promise<SqlResult>,
    ctx: PluginContext
  ): Promise<SqlResult> {
    try {
      return await fn();
    } catch (e) {
      await DialogUtils.progressBar?.end();

      if (!(e instanceof Error || e instanceof SystemError || e instanceof UserError)) {
        e = new Error(e.toString());
      }
      if (!(e instanceof SystemError) && !(e instanceof UserError)) {
        ctx.logProvider?.error(e.message);
      }

      let res: SqlResult;
      if (e instanceof SystemError || e instanceof UserError) {
        res = err(e);
      } else {
        res = err(
          SqlResultFactory.SystemError(
            ErrorMessage.UnhandledError.name,
            ErrorMessage.UnhandledError.message(),
            e
          )
        );
      }
      const errorCode = res.error.source + "." + res.error.name;
      const errorType =
        res.error instanceof SystemError ? Telemetry.systemError : Telemetry.userError;
      TelemetryUtils.init(ctx);
      let errorMessage = res.error.message;
      if (res.error.innerError) {
        errorMessage += ` Detailed error: ${e.innerError.message}.`;
      }
      TelemetryUtils.sendErrorEvent(stage, errorCode, errorType, errorMessage);
      return res;
    }
  }
}

export default new SqlPlugin();
