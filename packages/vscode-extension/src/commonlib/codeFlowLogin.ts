/* eslint-disable @typescript-eslint/ban-types */
/* eslint-disable @typescript-eslint/no-var-requires */
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { PublicClientApplication, AccountInfo, Configuration, TokenCache } from "@azure/msal-node";
import * as express from "express";
import * as http from "http";
import * as fs from "fs-extra";
import * as path from "path";
import { Mutex } from "async-mutex";
import { UserError, SystemError, returnUserError } from "@microsoft/teamsfx-api";
import VsCodeLogInstance from "./log";
import * as crypto from "crypto";
import { AddressInfo } from "net";
import { loadAccountId, saveAccountId, UTF8 } from "./cacheAccess";
import * as stringUtil from "util";
import * as StringResources from "../resources/Strings.json";
import { loggedIn, loggedOut, loggingIn } from "./common/constant";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";

interface Deferred<T> {
  resolve: (result: T | Promise<T>) => void;
  reject: (reason: any) => void;
}

export class CodeFlowLogin {
  pca: PublicClientApplication | undefined;
  account: AccountInfo | undefined;
  scopes: string[] | undefined;
  config: Configuration | undefined;
  port: number | undefined;
  mutex: Mutex | undefined;
  msalTokenCache: TokenCache | undefined;
  accountName: string;
  status: string | undefined;

  constructor(scopes: string[], config: Configuration, port: number, accountName: string) {
    this.scopes = scopes;
    this.config = config;
    this.port = port;
    this.mutex = new Mutex();
    this.pca = new PublicClientApplication(this.config!);
    this.msalTokenCache = this.pca.getTokenCache();
    this.accountName = accountName;
    this.status = loggedOut;
  }

  async reloadCache() {
    const accountCache = await loadAccountId(this.accountName);
    if (accountCache) {
      const dataCache = await this.msalTokenCache!.getAccountByHomeId(accountCache);
      if (dataCache) {
        this.account = dataCache;
        this.status = loggedIn;
      }
    }
  }

  async login(): Promise<string> {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.LoginStart, {
      [TelemetryProperty.AccountType]: this.accountName,
    });
    const codeVerifier = CodeFlowLogin.toBase64UrlEncoding(
      crypto.randomBytes(32).toString("base64")
    );
    const codeChallenge = CodeFlowLogin.toBase64UrlEncoding(
      await CodeFlowLogin.sha256(codeVerifier)
    );
    let serverPort = this.port;

    // try get an unused port
    const app = express();
    const server = app.listen(serverPort);
    serverPort = (server.address() as AddressInfo).port;

    const authCodeUrlParameters = {
      scopes: this.scopes!,
      codeChallenge: codeChallenge,
      codeChallengeMethod: "S256",
      redirectUri: `http://localhost:${serverPort}`,
      prompt: "select_account",
    };

    let deferredRedirect: Deferred<string>;
    const redirectPromise: Promise<string> = new Promise<string>(
      (resolve, reject) => (deferredRedirect = { resolve, reject })
    );

    app.get("/", (req: express.Request, res: express.Response) => {
      this.status = loggingIn;
      const tokenRequest = {
        code: req.query.code as string,
        scopes: this.scopes!,
        redirectUri: `http://localhost:${serverPort}`,
        codeVerifier: codeVerifier,
      };

      this.pca!.acquireTokenByCode(tokenRequest)
        .then(async (response) => {
          if (response) {
            if (response.account) {
              await this.mutex?.runExclusive(async () => {
                this.account = response.account!;
                this.status = loggedIn;
                await saveAccountId(this.accountName, this.account.homeAccountId);
              });
              deferredRedirect.resolve(response.accessToken);

              const resultFilePath = path.join(__dirname, "./codeFlowResult/index.html");
              if (fs.existsSync(resultFilePath)) {
                sendFile(res, resultFilePath, "text/html; charset=utf-8");
              } else {
                // do not break if result file has issue
                VsCodeLogInstance.error(
                  "[Login] " + StringResources.vsc.codeFlowLogin.resultFileNotFound
                );
                res.sendStatus(200);
              }
            }
          } else {
            throw new Error("get no response");
          }
        })
        .catch((error) => {
          this.status = loggedOut;
          VsCodeLogInstance.error("[Login] " + error.message);
          deferredRedirect.reject(error);
          res.status(500).send(error);
        });
    });

    const codeTimer = setTimeout(() => {
      if (this.account) {
        this.status = loggedIn;
      } else {
        this.status = loggedOut;
      }
      deferredRedirect.reject(
        returnUserError(
          new Error(StringResources.vsc.codeFlowLogin.loginTimeoutDescription),
          StringResources.vsc.codeFlowLogin.loginComponent,
          StringResources.vsc.codeFlowLogin.loginTimeoutTitle
        )
      );
    }, 5 * 60 * 1000); // keep the same as azure login

    function cancelCodeTimer() {
      clearTimeout(codeTimer);
    }

    let accessToken = undefined;
    try {
      await this.startServer(server, serverPort!);
      this.pca!.getAuthCodeUrl(authCodeUrlParameters).then(async (response: string) => {
        vscode.env.openExternal(vscode.Uri.parse(response));
      });

      redirectPromise.then(cancelCodeTimer, cancelCodeTimer);
      accessToken = await redirectPromise;
    } catch (e) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Login, {
        [TelemetryProperty.AccountType]: this.accountName,
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.UserId]: "",
        [TelemetryProperty.Internal]: "false",
        [TelemetryProperty.ErrorType]:
          e instanceof UserError ? TelemetryErrorType.UserError : TelemetryErrorType.SystemError,
        [TelemetryProperty.ErrorCode]: `${e.source}.${e.name}`,
        [TelemetryProperty.ErrorMessage]: `${e.message}`,
      });
      throw e;
    } finally {
      if (accessToken) {
        const tokenJson = ConvertTokenToJson(accessToken);
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Login, {
          [TelemetryProperty.AccountType]: this.accountName,
          [TelemetryProperty.Success]: TelemetrySuccess.Yes,
          [TelemetryProperty.UserId]: (tokenJson as any).oid ? (tokenJson as any).oid : "",
          [TelemetryProperty.Internal]: (tokenJson as any).upn?.endsWith("@microsoft.com")
            ? "true"
            : "false",
        });
      }
      server.close();
    }

    return accessToken;
  }

  async logout(): Promise<boolean> {
    try {
      const accountCache = await loadAccountId(this.accountName);
      if (accountCache) {
        const dataCache = await this.msalTokenCache!.getAccountByHomeId(accountCache);
        if (dataCache) {
          this.msalTokenCache?.removeAccount(dataCache);
        }
      }

      await saveAccountId(this.accountName, undefined);
      this.account = undefined;
      this.status = loggedOut;
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.SignOut, {
        [TelemetryProperty.AccountType]: this.accountName,
        [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      });
      return true;
    } catch (e) {
      VsCodeLogInstance.error("[Logout " + this.accountName + "] " + e.message);
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.SignOut, e, {
        [TelemetryProperty.AccountType]: this.accountName,
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.ErrorType]:
          e instanceof UserError ? TelemetryErrorType.UserError : TelemetryErrorType.SystemError,
        [TelemetryProperty.ErrorCode]: `${e.source}.${e.name}`,
        [TelemetryProperty.ErrorMessage]: `${e.message}`,
      });
      return false;
    }
  }

  async getToken(refresh = true): Promise<string | undefined> {
    try {
      if (!this.account) {
        const accessToken = await this.login();
        return accessToken;
      } else {
        return this.pca!.acquireTokenSilent({
          account: this.account,
          scopes: this.scopes!,
          forceRefresh: false,
        })
          .then((response) => {
            if (response) {
              return response.accessToken;
            } else {
              return undefined;
            }
          })
          .catch(async (error) => {
            VsCodeLogInstance.error(
              "[Login] " +
                stringUtil.format(
                  StringResources.vsc.codeFlowLogin.silentAcquireToken,
                  error.message
                )
            );
            await this.logout();
            if (refresh) {
              const accessToken = await this.login();
              return accessToken;
            }
            return undefined;
          });
      }
    } catch (error) {
      VsCodeLogInstance.error("[Login] " + error.message);
      if (
        error.name !== StringResources.vsc.codeFlowLogin.loginTimeoutTitle &&
        error.name !== StringResources.vsc.codeFlowLogin.loginPortConflictTitle
      ) {
        throw LoginCodeFlowError(error);
      } else {
        throw error;
      }
    }
  }

  async startServer(server: http.Server, port: number): Promise<string> {
    // handle port timeout
    let defferedPort: Deferred<string>;
    const portPromise: Promise<string> = new Promise<string>(
      (resolve, reject) => (defferedPort = { resolve, reject })
    );
    const portTimer = setTimeout(() => {
      defferedPort.reject(
        returnUserError(
          new Error(StringResources.vsc.codeFlowLogin.loginPortConflictDescription),
          StringResources.vsc.codeFlowLogin.loginComponent,
          StringResources.vsc.codeFlowLogin.loginPortConflictTitle
        )
      );
    }, 5000);

    function cancelPortTimer() {
      clearTimeout(portTimer);
    }

    server.on("listening", () => {
      defferedPort.resolve(`Code login server listening on port ${port}`);
    });
    portPromise.then(cancelPortTimer, cancelPortTimer);
    return portPromise;
  }

  static toBase64UrlEncoding(base64string: string) {
    return base64string.replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");
  }

  static sha256(s: string | Uint8Array): Promise<string> {
    return require("crypto").createHash("sha256").update(s).digest("base64");
  }
}

function sendFile(res: http.ServerResponse, filepath: string, contentType: string) {
  fs.readFile(filepath, (err, body) => {
    if (err) {
      VsCodeLogInstance.error(err.message);
    } else {
      res.writeHead(200, {
        "Content-Length": body.length,
        "Content-Type": contentType,
      });
      res.end(body);
    }
  });
}

export function LoginFailureError(innerError?: any): UserError {
  return new UserError(
    StringResources.vsc.codeFlowLogin.loginFailureTitle,
    StringResources.vsc.codeFlowLogin.loginFailureDescription,
    "Login",
    new Error().stack,
    undefined,
    innerError
  );
}

export function LoginCodeFlowError(innerError?: any): SystemError {
  return new SystemError(
    StringResources.vsc.codeFlowLogin.loginCodeFlowFailureTitle,
    StringResources.vsc.codeFlowLogin.loginCodeFlowFailureDescription,
    StringResources.vsc.codeFlowLogin.loginComponent,
    new Error().stack,
    undefined,
    innerError
  );
}

export function ConvertTokenToJson(token: string): object {
  const array = token!.split(".");
  const buff = Buffer.from(array[1], "base64");
  return JSON.parse(buff.toString(UTF8));
}
