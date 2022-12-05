// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { DriverContext } from "../interface/commonArgs";
import { Service } from "typedi";
import { Constants } from "./constant";
import { deployArgs } from "./interface";
import { ArmDeployImpl } from "./deployImpl";
import { FxError, Result } from "@microsoft/teamsfx-api";
import { WrapDriverContext, wrapRun } from "../util/wrapUtil";
import { getLocalizedString } from "../../../common/localizeUtils";

@Service(Constants.actionName) // DO NOT MODIFY the service name
export class ArmDeployDriver implements StepDriver {
  description = getLocalizedString("driver.arm.description.deploy");
  public async run(
    args: deployArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(context, Constants.actionName, Constants.actionName);
    const impl = new ArmDeployImpl(args, wrapContext);
    const wrapRes = await wrapRun(wrapContext, () => impl.run());
    return wrapRes as Result<Map<string, string>, FxError>;
  }

  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(ctx, Constants.actionName, Constants.actionName);
    const impl = new ArmDeployImpl(args as deployArgs, wrapContext);
    const wrapRes = await wrapRun(wrapContext, () => impl.run(), true);
    return wrapRes as ExecutionResult;
  }
}
