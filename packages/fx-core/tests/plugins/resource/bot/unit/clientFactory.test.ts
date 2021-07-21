// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import "mocha";
import * as chai from "chai";

import {
  createAzureBotServiceClient,
  createWebSiteMgmtClient,
  ensureResourceProvider,
} from "../../../../../src/plugins/resource/bot/clientFactory";
import { generateFakeServiceClientCredentials } from "./utils";
import { AzureBotService } from "@azure/arm-botservice";
import { WebSiteManagementClient } from "@azure/arm-appservice";
import { PluginError } from "../../../../../src/plugins/resource/bot/errors";
import { Messages } from "./messages";

describe("Client Factory", () => {
  describe("createAzureBotServiceClient", () => {
    it("Happy Path", () => {
      // Arrange
      const subsName = "subsName";
      const credentials = generateFakeServiceClientCredentials();

      // Act
      const client = createAzureBotServiceClient(credentials, subsName);

      // Assert
      chai.assert.isTrue(client instanceof AzureBotService);
    });

    it("Empty Subscription", () => {
      // Arrange
      const subsName = "";
      const credentials = generateFakeServiceClientCredentials();

      // Act
      try {
        createAzureBotServiceClient(credentials, subsName);
      } catch (e) {
        chai.assert.isTrue(e instanceof PluginError);
        return;
      }

      chai.assert.fail(Messages.ShouldNotReachHere);
    });
  });

  describe("createWebSiteMgmtClient", () => {
    it("Happy Path", () => {
      // Arrange
      const subsName = "subsName";
      const credentials = generateFakeServiceClientCredentials();

      // Act
      const client = createWebSiteMgmtClient(credentials, subsName);

      // Assert
      chai.assert.isTrue(client instanceof WebSiteManagementClient);
    });

    it("Empty Subscription", () => {
      // Arrange
      const subsName = "";
      const credentials = generateFakeServiceClientCredentials();

      // Act
      try {
        createWebSiteMgmtClient(credentials, subsName);
      } catch (e) {
        chai.assert.isTrue(e instanceof PluginError);
        return;
      }

      chai.assert.fail(Messages.ShouldNotReachHere);
    });
  });

  describe("create", () => {
    it("Test ensureResourceProvider with existence", async () => {
      // Arrange
      const item: any = { registrationState: "Registered" };
      const namespace = ["ut"];
      const credentials = generateFakeServiceClientCredentials();
      const client: any = {
        get: (namespace: string) => item,
        register: (namespace: string) => item,
      };

      // Act
      const res = await ensureResourceProvider(client, namespace);

      // Assert
      chai.assert.deepEqual(res, [item]);
    });

    it("Test ensureResourceProvider", async () => {
      // Arrange
      let item: any = { registrationState: "Unregistered" };
      const namespace = ["ut"];
      const client: any = {
        get: (namespace: string) => item,
        register: (namespace: string) => {
          item = {};
          item = {...item, $namespace: {"registrationState": "Registered"}};
          return item;
        },
      };

      // Act
      const res = await ensureResourceProvider(client, namespace);

      // Assert
      chai.assert.deepEqual(res, [item]);
    });
  });
});
