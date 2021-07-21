// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { ChildProcess, spawn } from "child_process";
import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import treeKill from "tree-kill";
import { ServiceLogWriter } from "./serviceLogWriter";

interface TaskOptions {
  shell: boolean | string;
  cwd?: string;
  env?: NodeJS.ProcessEnv;
}

export interface TaskResult {
  command: string;
  args?: string[];
  options?: TaskOptions;
  success: boolean;
  stdout: string[];
  stderr: string[];
  exitCode: number | null;
}

export class Task {
  private taskTitle: string;
  private background: boolean;
  private command: string;
  private args?: string[];
  private options?: TaskOptions;

  private resolved = false;
  private task: ChildProcess | undefined;

  constructor(
    taskTitle: string,
    background: boolean,
    command: string,
    args?: string[],
    options?: TaskOptions
  ) {
    this.taskTitle = taskTitle;
    this.background = background;
    this.command = command;
    this.args = args;
    this.options = options;
  }

  /**
   * wait for the task to end
   */
  public async wait(
    startCallback: (taskTitle: string, background: boolean) => Promise<void>,
    stopCallback: (
      taskTitle: string,
      background: boolean,
      result: TaskResult
    ) => Promise<FxError | null>
  ): Promise<Result<TaskResult, FxError>> {
    await startCallback(this.taskTitle, this.background);
    this.task = spawn(this.command, this.args, this.options);
    const stdout: string[] = [];
    const stderr: string[] = [];
    return new Promise((resolve) => {
      this.task?.stdout?.on("data", (data) => {
        // TODO: log
        stdout.push(data.toString());
      });
      this.task?.stderr?.on("data", (data) => {
        // TODO: log
        stderr.push(data.toString());
      });
      this.task?.on("exit", async () => {
        const result: TaskResult = {
          command: this.command,
          options: this.options,
          success: this.task?.exitCode === 0,
          stdout: stdout,
          stderr: stderr,
          exitCode: this.task?.exitCode === undefined ? null : this.task?.exitCode,
        };
        const error = await stopCallback(this.taskTitle, this.background, result);
        if (error) {
          resolve(err(error));
        } else {
          resolve(ok(result));
        }
      });
    });
  }

  /**
   * wait until stdout of the task matches the pattern or the task ends
   */
  public async waitFor(
    pattern: RegExp,
    startCallback: (taskTitle: string, background: boolean) => Promise<void>,
    stopCallback: (
      taskTitle: string,
      background: boolean,
      result: TaskResult,
      serviceLogWriter?: ServiceLogWriter
    ) => Promise<FxError | null>,
    serviceLogWriter?: ServiceLogWriter
  ): Promise<Result<TaskResult, FxError>> {
    await startCallback(this.taskTitle, this.background);
    this.task = spawn(this.command, this.args, this.options);
    const stdout: string[] = [];
    const stderr: string[] = [];
    return new Promise((resolve) => {
      this.task?.stdout?.on("data", async (data) => {
        const dataStr = data.toString();
        await serviceLogWriter?.write(this.taskTitle, dataStr);
        stdout.push(dataStr);
        if (!this.resolved) {
          const match = pattern.test(dataStr);
          if (match) {
            this.resolved = true;
            const result: TaskResult = {
              command: this.command,
              options: this.options,
              success: true,
              stdout: stdout,
              stderr: stderr,
              exitCode: null,
            };
            const error = await stopCallback(
              this.taskTitle,
              this.background,
              result,
              serviceLogWriter
            );
            if (error) {
              resolve(err(error));
            } else {
              resolve(ok(result));
            }
          }
        }
      });
      this.task?.stderr?.on("data", async (data) => {
        const dataStr = data.toString();
        await serviceLogWriter?.write(this.taskTitle, dataStr);
        stderr.push(dataStr);
      });

      this.task?.on("exit", async () => {
        if (!this.resolved) {
          this.resolved = true;
          const result: TaskResult = {
            command: this.command,
            options: this.options,
            success: false,
            stdout: stdout,
            stderr: stderr,
            exitCode: this.task?.exitCode === undefined ? null : this.task?.exitCode,
          };
          const error = await stopCallback(
            this.taskTitle,
            this.background,
            result,
            serviceLogWriter
          );
          if (error) {
            resolve(err(error));
          } else {
            resolve(ok(result));
          }
        }
      });
    });
  }

  public async terminate(): Promise<void> {
    return new Promise((resolve) => {
      if (this.task?.exitCode) {
        resolve();
      }
      const pid = this.task?.pid;
      if (pid === undefined) {
        resolve();
      } else {
        treeKill(pid, (error) => {
          if (error) {
            // ignore any error
            resolve();
          } else {
            resolve();
          }
        });
      }
    });
  }
}
