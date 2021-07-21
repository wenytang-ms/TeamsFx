// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance } from "axios";
import { SystemError, LogProvider } from "@microsoft/teamsfx-api";
import { IAppDefinition } from "./interfaces/IAppDefinition";
import { AppStudioError } from "./errors";
import { IPublishingAppDenition } from "./interfaces/IPublishingAppDefinition";
import { AppStudioResultFactory } from "./results";

// eslint-disable-next-line @typescript-eslint/no-namespace
export namespace AppStudioClient {
  type Icon = {
    type: "color" | "outline" | "sharePointPreviewImage";
    name: "color" | "outline" | "sharePointPreviewImage";
    base64String: string;
  };

  const baseUrl = "https://dev.teams.microsoft.com";

  /**
   * Creates a new axios instance to call app studio to prevent setting the accessToken on global instance.
   * @param {string}  appStudioToken
   * @returns {AxiosInstance}
   */
  function createRequesterWithToken(appStudioToken: string): AxiosInstance {
    const instance = axios.create({
      baseURL: baseUrl,
    });
    instance.defaults.headers.common["Authorization"] = `Bearer ${appStudioToken}`;
    return instance;
  }

  /**
   * Creates an app registration in app studio with the given configuration and returns the Teams app id.
   * @param {IAppDefinition}  appDefinition
   * @param {string}  appStudioToken
   * @param {LogProvider} logProvider
   * @param {string} colorIconContent - base64 encoded
   * @param {string} outlineIconContent - base64 encoded
   * @returns {Promise<IAppDefinition | undefined>}
   */
  export async function createApp(
    appDefinition: IAppDefinition,
    appStudioToken: string,
    logProvider?: LogProvider,
    colorIconContent?: string,
    outlineIconContent?: string
  ): Promise<IAppDefinition | undefined> {
    if (appDefinition && appStudioToken) {
      try {
        const requester = createRequesterWithToken(appStudioToken);
        const appDef = {
          ...appDefinition,
        };
        // /api/appdefinitions/import accepts icons as Base64-encoded strings.
        if (colorIconContent) {
          appDef.colorIcon = colorIconContent;
        }
        if (outlineIconContent) {
          appDef.outlineIcon = outlineIconContent;
        }
        const response = await requester.post(`/api/appdefinitions/import`, appDef);
        if (response && response.data) {
          const app = <IAppDefinition>response.data;
          await logProvider?.debug(`recieved data from app studio ${JSON.stringify(app)}`);

          if (app) {
            return app;
          }
        }
      } catch (e) {
        if (e instanceof Error) {
          await logProvider?.warning(`failed to create app due to ${e.name}: ${e.message}`);
        }
        return undefined;
      }
    }

    await logProvider?.warning(`invalid appDefinition or appStudioToken`);
    return undefined;
  }

  async function uploadIcon(
    teamsAppId: string,
    appStudioToken: string,
    colorIconContent: string,
    outlineIconContent: string,
    requester: AxiosInstance,
    logProvider?: LogProvider
  ): Promise<{ colorIconUrl: string; outlineIconUrl: string }> {
    await logProvider?.info(`uploading icon for teams ${teamsAppId}`);
    if (teamsAppId && appStudioToken) {
      try {
        const colorIcon: Icon = {
          name: "color",
          type: "color",
          base64String: colorIconContent,
        };
        const outlineIcon: Icon = {
          name: "outline",
          type: "outline",
          base64String: outlineIconContent,
        };
        const colorIconResult = requester.post(
          `/api/appdefinitions/${teamsAppId}/image`,
          colorIcon
        );
        const outlineIconResult = requester.post(
          `/api/appdefinitions/${teamsAppId}/image`,
          outlineIcon
        );
        const results = await Promise.all([colorIconResult, outlineIconResult]);
        await logProvider?.info(`successfully uploaded two icons`);
        return { colorIconUrl: results[0].data, outlineIconUrl: results[1].data };
      } catch (e) {
        if (e instanceof Error) {
          await logProvider?.warning(`failed to upload icon due to ${e.name}: ${e.message}`);
        }
        throw e;
      }
    }
    throw new Error(`teamsAppId or appStudioToken is invalid`);
  }

  export async function getApp(
    teamsAppId: string,
    appStudioToken: string,
    logProvider?: LogProvider
  ): Promise<IAppDefinition> {
    const requester = createRequesterWithToken(appStudioToken);
    try {
      const response = await requester.get(`/api/appdefinitions/${teamsAppId}`);
      if (response && response.data) {
        const app = <IAppDefinition>response.data;
        if (app && app.teamsAppId && app.teamsAppId === teamsAppId) {
          return app;
        } else {
          await logProvider?.error(
            `teamsAppId mismatch. Input: ${teamsAppId}. Got: ${app.teamsAppId}`
          );
        }
      }
    } catch (e) {
      if (e instanceof Error) {
        await logProvider?.warning(
          `Cannot get the app definition with app ID ${teamsAppId}, due to ${e.name}: ${e.message}`
        );
      }
      throw new Error(
        `Cannot get the app definition with app ID ${teamsAppId}, due to ${e.name}: ${e.message}`
      );
    }
    throw new Error(`Cannot get the app definition with app ID ${teamsAppId}`);
  }

  /**
   * Updates an existing app if it exists with the configuration given.  Returns whether or not it was successful.
   * @param {string}  teamsAppId
   * @param {IAppDefinition} appDefinition
   * @param {string}  appStudioToken
   * @param {LogProvider} logProvider
   * @param {string} colorIconContent - base64 encoded
   * @param {string} outlineIconContent - base64 encoded
   * @returns {Promise<IAppDefinition>}
   */
  export async function updateApp(
    teamsAppId: string,
    appDefinition: IAppDefinition,
    appStudioToken: string,
    logProvider?: LogProvider,
    colorIconContent?: string,
    outlineIconContent?: string
  ): Promise<IAppDefinition> {
    if (appDefinition && appStudioToken) {
      try {
        const requester = createRequesterWithToken(appStudioToken);

        // Get userlist from existing app
        const existingAppDefinition = await getApp(teamsAppId, appStudioToken, logProvider);
        const userlist = existingAppDefinition.userList;
        appDefinition.userList = userlist;

        let result: { colorIconUrl: string; outlineIconUrl: string } | undefined;
        if (colorIconContent && outlineIconContent) {
          result = await uploadIcon(
            teamsAppId,
            appStudioToken,
            colorIconContent,
            outlineIconContent,
            requester,
            logProvider
          );
          if (!result) {
            await logProvider?.error(`failed to upload color icon for: ${teamsAppId}`);
            throw new Error(`failed to upload icons for ${teamsAppId}`);
          }
          appDefinition.colorIcon = result.colorIconUrl;
          appDefinition.outlineIcon = result.outlineIconUrl;
        }
        const response = await requester.post(
          `/api/appdefinitions/${teamsAppId}/override`,
          appDefinition
        );
        if (response && response.data) {
          const app = <IAppDefinition>response.data;

          if (app && app.teamsAppId && app.teamsAppId === teamsAppId) {
            return app;
          } else {
            await logProvider?.error(
              `teamsAppId mismatch. Input: ${teamsAppId}. Got: ${app.teamsAppId}`
            );
          }
        }
      } catch (e) {
        if (e instanceof Error) {
          await logProvider?.warning(`failed to update app due to ${e.name}: ${e.message}`);
        }
        throw new Error(`failed to update app due to ${e.name}: ${e.message}`);
      }
    }

    throw new Error(`invalid appDefinition[${appDefinition}] or appStudioToken[${appStudioToken}]`);
  }

  export async function validateManifest(
    manifestString: string,
    appStudioToken: string
  ): Promise<string[]> {
    try {
      const requester = createRequesterWithToken(appStudioToken);
      const buffer = Buffer.from(manifestString, "utf8");
      const response = await requester.post("/api/appdefinitions/prevalidation", buffer, {
        headers: { "Content-Type": "application/json" },
      });
      if (response && response.data) {
        let validationResult: string[] = [];
        validationResult = validationResult.concat(response.data.errors);
        validationResult = validationResult.concat(response.data.warnings);
        validationResult = validationResult.concat(response.data.info);
        return validationResult;
      } else {
        throw AppStudioResultFactory.SystemError(
          AppStudioError.ValidationFailedError.name,
          AppStudioError.ValidationFailedError.message(["Unknown reason"])
        );
      }
    } catch (e) {
      throw AppStudioResultFactory.SystemError(
        AppStudioError.ValidationFailedError.name,
        AppStudioError.ValidationFailedError.message(["Unknown reason"]),
        e
      );
    }
  }

  /**
   * Publish Teams app to Teams App Catalog
   * @param teamsAppId
   * @param file
   * @param appStudioToken
   * @returns
   */
  export async function publishTeamsApp(
    teamsAppId: string,
    file: Buffer,
    appStudioToken: string
  ): Promise<string> {
    try {
      const requester = createRequesterWithToken(appStudioToken);
      const response = await requester.post("/api/publishing", file, {
        headers: { "Content-Type": "application/zip" },
      });

      if (response && response.data) {
        if (response.data.error) {
          // To avoid App Studio BadGateway error
          // The app is actually published to app catalog.
          if (response.data.error.code === "BadGateway") {
            const appDefinition = await getAppByTeamsAppId(teamsAppId, appStudioToken);
            if (appDefinition) {
              return appDefinition.teamsAppId;
            }
          }
          throw AppStudioResultFactory.SystemError(
            AppStudioError.TeamsAppPublishFailedError.name,
            AppStudioError.TeamsAppPublishFailedError.message(teamsAppId),
            `code: ${response.data.error.code}, message: ${response.data.error.message}`
          );
        } else {
          return response.data.id;
        }
      } else {
        throw AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppPublishFailedError.name,
          AppStudioError.TeamsAppPublishFailedError.message(teamsAppId)
        );
      }
    } catch (error) {
      if (error instanceof SystemError) {
        throw error;
      } else {
        throw AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppPublishFailedError.name,
          AppStudioError.TeamsAppPublishFailedError.message(teamsAppId),
          error
        );
      }
    }
  }

  /**
   * Update existed publish request
   * @param teamsAppId
   * @param file
   * @param appStudioToken
   * @returns
   */
  export async function publishTeamsAppUpdate(
    teamsAppId: string,
    file: Buffer,
    appStudioToken: string
  ): Promise<string> {
    try {
      // Get App Definition from Teams App Catalog
      const appDefinition = await getAppByTeamsAppId(teamsAppId, appStudioToken);

      const requester = createRequesterWithToken(appStudioToken);
      let response = null;
      if (appDefinition) {
        // update the existing app
        response = await requester.post(
          `/api/publishing/${appDefinition.teamsAppId}/appdefinitions`,
          file,
          { headers: { "Content-Type": "application/zip" } }
        );
      } else {
        throw AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppPublishFailedError.name,
          AppStudioError.TeamsAppPublishFailedError.message(teamsAppId)
        );
      }

      if (response && response.data) {
        if (response.data.error || response.data.errorMessage) {
          throw AppStudioResultFactory.SystemError(
            AppStudioError.TeamsAppPublishFailedError.name,
            AppStudioError.TeamsAppPublishFailedError.message(teamsAppId),
            response.data.error?.message || response.data.errorMessage
          );
        } else {
          return response.data.id;
        }
      } else {
        throw AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppPublishFailedError.name,
          AppStudioError.TeamsAppPublishFailedError.message(teamsAppId)
        );
      }
    } catch (error) {
      if (error instanceof SystemError) {
        throw error;
      } else {
        throw AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppPublishFailedError.name,
          AppStudioError.TeamsAppPublishFailedError.message(teamsAppId),
          error
        );
      }
    }
  }

  export async function getAppByTeamsAppId(
    teamsAppId: string,
    appStudioToken: string
  ): Promise<IPublishingAppDenition | undefined> {
    const requester = createRequesterWithToken(appStudioToken);
    const response = await requester.get(`/api/publishing/${teamsAppId}`);
    if (response && response.data && response.data.value && response.data.value.length > 0) {
      const appdefinitions: IPublishingAppDenition[] = response.data.value[0].appDefinitions.map(
        (item: any) => {
          return {
            lastModifiedDateTime: item.lastModifiedDateTime
              ? new Date(item.lastModifiedDateTime)
              : null,
            publishingState: item.publishingState,
            teamsAppId: item.teamsAppId,
            displayName: item.displayName,
          };
        }
      );
      return appdefinitions[appdefinitions.length - 1];
    } else {
      return undefined;
    }
  }
}
