// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { PublicClientApplication, AccountInfo, Configuration, TokenCache } from "@azure/msal-node";
import express from "express";
import * as http from "http";
import * as fs from "fs-extra";
import * as path from "path";
import { Mutex } from "async-mutex";
import {
  returnSystemError,
  returnUserError,
  SystemError,
  UserError,
  LogLevel,
  Colors,
} from "@microsoft/teamsfx-api";
import CliCodeLogInstance from "./log";
import * as crypto from "crypto";
import { AddressInfo } from "net";
import { loadAccountId, saveAccountId, UTF8 } from "./cacheAccess";
import open from "open";
import { azureLoginMessage, env, m365LoginMessage, MFACode } from "./common/constant";
import * as constants from "../constants";
import CliTelemetry from "../telemetry/cliTelemetry";
import {
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import { getColorizedString } from "../utils";

class ErrorMessage {
  static readonly loginFailureTitle = "LoginFail";
  static readonly loginFailureDescription =
    "Cannot retrieve user login information. Login with another account.";
  static readonly loginCodeFlowFailureTitle = "LoginCodeFail";
  static readonly loginCodeFlowFailureDescription =
    "Cannot get login code for token exchange. Login with another account.";
  static readonly loginTimeoutTitle = "LoginTimeout";
  static readonly loginTimeoutDescription = "Timeout waiting for login. Try again.";
  static readonly loginPortConflictTitle = "LoginPortConflict";
  static readonly loginPortConflictDescription = "Timeout waiting for port. Try again.";
  static readonly loginComponent = "login";
}

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
  socketMap: Map<number, any>;

  constructor(scopes: string[], config: Configuration, port: number, accountName: string) {
    this.scopes = scopes;
    this.config = config;
    this.port = port;
    this.mutex = new Mutex();
    this.pca = new PublicClientApplication(this.config!);
    this.msalTokenCache = this.pca.getTokenCache();
    this.accountName = accountName;
    this.socketMap = new Map();
  }

  async reloadCache() {
    const accountCache = await loadAccountId(this.accountName);
    if (accountCache) {
      const dataCache = await this.msalTokenCache!.getAccountByHomeId(accountCache);
      if (dataCache) {
        this.account = dataCache;
      }
    }
  }

  async login(): Promise<string> {
    CliTelemetry.sendTelemetryEvent(TelemetryEvent.AccountLoginStart, {
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
    let lastSocketKey = 0;
    server.on("connection", (socket) => {
      const socketKey = ++lastSocketKey;
      this.socketMap.set(socketKey, socket);
      socket.on("close", () => {
        this.socketMap.delete(socketKey);
      });
    });

    server.on("close", () => {
      this.destroySockets();
    });

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
                await saveAccountId(this.accountName, this.account.homeAccountId);
              });
              deferredRedirect.resolve(response.accessToken);

              sendFile(
                res,
                path.join(__dirname, "./codeFlowResult/index.html"),
                "text/html; charset=utf-8",
                this.accountName!
              );
            }
          } else {
            throw new Error("get no response");
          }
        })
        .catch((error) => {
          CliCodeLogInstance.necessaryLog(LogLevel.Error, "[Login] " + error.message);
          deferredRedirect.reject(error);
          res.status(500).send(error);
        });
    });

    const codeTimer = setTimeout(() => {
      deferredRedirect.reject(
        returnUserError(
          new Error(ErrorMessage.loginTimeoutDescription),
          ErrorMessage.loginComponent,
          ErrorMessage.loginTimeoutTitle
        )
      );
    }, 5 * 60 * 1000);

    function cancelCodeTimer() {
      clearTimeout(codeTimer);
    }

    let accessToken = undefined;
    try {
      await this.startServer(server, serverPort!);
      this.pca!.getAuthCodeUrl(authCodeUrlParameters).then(async (url: string) => {
        if (this.accountName == "azure") {
          const message = [
            {
              content: `[${constants.cliSource}] ${azureLoginMessage}`,
              color: Colors.BRIGHT_WHITE,
            },
            { content: url, color: Colors.BRIGHT_CYAN },
          ];
          CliCodeLogInstance.necessaryLog(LogLevel.Info, getColorizedString(message));
        } else {
          const message = [
            { content: `[${constants.cliSource}] ${m365LoginMessage}`, color: Colors.BRIGHT_WHITE },
            { content: url, color: Colors.BRIGHT_CYAN },
          ];
          CliCodeLogInstance.necessaryLog(LogLevel.Info, getColorizedString(message));
        }
        open(url);
      });

      redirectPromise.then(cancelCodeTimer, cancelCodeTimer);
      accessToken = await redirectPromise;
    } catch (e) {
      CliTelemetry.sendTelemetryEvent(TelemetryEvent.AccountLogin, {
        [TelemetryProperty.AccountType]: this.accountName,
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.UserId]: "",
        [TelemetryProperty.Internal]: "",
        [TelemetryProperty.ErrorType]:
          e instanceof UserError ? TelemetryErrorType.UserError : TelemetryErrorType.SystemError,
        [TelemetryProperty.ErrorCode]: `${e.source}.${e.name}`,
        [TelemetryProperty.ErrorMessage]: `${e.message}`,
      });
      throw e;
    } finally {
      if (accessToken) {
        const tokenJson = ConvertTokenToJson(accessToken);
        CliTelemetry.sendTelemetryEvent(TelemetryEvent.AccountLogin, {
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
    const accountCache = await loadAccountId(this.accountName);
    if (accountCache) {
      const dataCache = await this.msalTokenCache!.getAccountByHomeId(accountCache);
      if (dataCache) {
        this.msalTokenCache?.removeAccount(dataCache);
      }
    }

    await saveAccountId(this.accountName, undefined);
    return true;
  }

  async getToken(refresh = true): Promise<string | undefined> {
    try {
      if (!this.account) {
        await this.reloadCache();
      }
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
            CliCodeLogInstance.necessaryLog(
              LogLevel.Error,
              "[Login] silent acquire token : " + error.message
            );
            await this.logout();
            if (refresh) {
              const accessToken = await this.login();
              return accessToken;
            } else {
              return undefined;
            }
          });
      }
    } catch (error) {
      CliCodeLogInstance.necessaryLog(LogLevel.Error, "[Login] " + error.message);
      if (
        error.name !== ErrorMessage.loginTimeoutTitle &&
        error.name !== ErrorMessage.loginPortConflictTitle
      ) {
        throw LoginCodeFlowError(error);
      } else {
        throw error;
      }
    }
  }

  async getTenantToken(tenantId: string): Promise<string | undefined> {
    try {
      if (!this.account) {
        await this.reloadCache();
      }
      if (this.account) {
        return this.pca!.acquireTokenSilent({
          authority: env.activeDirectoryEndpointUrl + tenantId,
          account: this.account,
          scopes: this.scopes!,
          forceRefresh: true,
        })
          .then((response) => {
            if (response) {
              return response.accessToken;
            } else {
              return undefined;
            }
          })
          .catch(async (error) => {
            if (error.message.indexOf(MFACode) >= 0) {
              throw error;
            } else {
              CliCodeLogInstance.necessaryLog(
                LogLevel.Error,
                "[Login] getTenantToken acquireTokenSilent : " + error.message
              );
              const accountList = await this.msalTokenCache?.getAllAccounts();
              for (let i = 0; i < accountList!.length; ++i) {
                this.msalTokenCache?.removeAccount(accountList![i]);
              }
              this.config!.auth.authority = env.activeDirectoryEndpointUrl + tenantId;
              this.pca = new PublicClientApplication(this.config!);
              const accessToken = await this.login();
              return accessToken;
            }
          });
      } else {
        return undefined;
      }
    } catch (error) {
      CliCodeLogInstance.necessaryLog(LogLevel.Error, "[Login] getTenantToken : " + error.message);
      throw LoginFailureError(error);
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
        returnSystemError(
          new Error(ErrorMessage.loginPortConflictDescription),
          ErrorMessage.loginComponent,
          ErrorMessage.loginPortConflictTitle
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

  destroySockets(): void {
    for (const key of this.socketMap.keys()) {
      this.socketMap.get(key).destroy();
    }
  }

  static toBase64UrlEncoding(base64string: string) {
    return base64string.replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");
  }

  static sha256(s: string | Uint8Array): Promise<string> {
    return new Promise((solve) => solve(crypto.createHash("sha256").update(s).digest("base64")));
  }
}

function sendFile(
  res: http.ServerResponse,
  filepath: string,
  contentType: string,
  accountName: string
) {
  fs.readFile(filepath, (err, body) => {
    if (err) {
      CliCodeLogInstance.necessaryLog(LogLevel.Error, err.message);
    } else {
      let data = body.toString();
      data = data.replace(/\${accountName}/g, accountName == "azure" ? "Azure" : "M365");
      body = Buffer.from(data, UTF8);
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
    ErrorMessage.loginCodeFlowFailureTitle,
    ErrorMessage.loginCodeFlowFailureDescription,
    ErrorMessage.loginComponent,
    new Error().stack,
    undefined,
    innerError
  );
}

export function LoginCodeFlowError(innerError?: any): SystemError {
  return new SystemError(
    ErrorMessage.loginCodeFlowFailureTitle,
    ErrorMessage.loginCodeFlowFailureDescription,
    ErrorMessage.loginComponent,
    new Error().stack,
    undefined,
    innerError
  );
}

export function ConvertTokenToJson(token: string): any {
  const array = token!.split(".");
  const buff = Buffer.from(array[1], "base64");
  return JSON.parse(buff.toString(UTF8));
}
