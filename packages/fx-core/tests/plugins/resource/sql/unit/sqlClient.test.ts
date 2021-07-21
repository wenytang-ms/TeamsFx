import "mocha";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { TestHelper } from "../helper";
import { SqlPlugin } from "../../../../../src/plugins/resource/sql";
import * as dotenv from "dotenv";
import { PluginContext, UserError } from "@microsoft/teamsfx-api";
import * as msRestNodeAuth from "@azure/ms-rest-nodeauth";
import * as faker from "faker";
import * as sinon from "sinon";
import { SqlClient } from "../../../../../src/plugins/resource/sql/sqlClient";
import { ErrorMessage } from "../../../../../src/plugins/resource/sql/errors";

chai.use(chaiAsPromised);

dotenv.config();

describe("sqlClient", () => {
  let sqlPlugin: SqlPlugin;
  let pluginContext: PluginContext;
  let credentials: msRestNodeAuth.TokenCredentialsBase;
  let client: SqlClient;

  before(async () => {
    credentials = new msRestNodeAuth.ApplicationTokenCredentials(
      faker.random.uuid(),
      faker.internet.url(),
      faker.internet.password()
    );
  });

  beforeEach(async () => {
    sqlPlugin = new SqlPlugin();
    pluginContext = await TestHelper.pluginContext(credentials);
    client = new SqlClient(pluginContext, sqlPlugin.sqlImpl.config);
  });

  afterEach(() => {
    sinon.restore();
  });

  it("existUser false", async function () {
    // Arrange
    sinon.stub(SqlClient.prototype, "doQuery").resolves([[{ value: 0 }]]);

    // Act
    const res = await client.existUser();

    // Assert
    chai.assert.isFalse(res);
  });

  it("existUser true", async function () {
    // Arrange
    sinon.stub(SqlClient.prototype, "doQuery").resolves([[{ value: 1 }]]);

    // Act
    const res = await client.existUser();

    // Assert
    chai.assert.isTrue(res);
  });

  it("existUser error", async function () {
    // Arrange
    sinon.stub(SqlClient.prototype, "doQuery").rejects(new Error("test error"));

    // Act
    try {
      await client.existUser();
    } catch (error) {
      // Assert
      chai.assert.include(error.message, "test error");
    }
  });

  it("addDatabaseUser error", async function () {
    // Arrange
    sinon
      .stub(SqlClient.prototype, "doQuery")
      .resolves()
      .onThirdCall()
      .rejects(new Error("test error"));

    // Act
    try {
      await client.addDatabaseUser();
    } catch (error) {
      // Assert
      chai.assert.include(error.message, ErrorMessage.GetDetail);
    }
  });

  it("addDatabaseUser admin error", async function () {
    // Arrange
    sinon
      .stub(SqlClient.prototype, "doQuery")
      .rejects(new Error("test error:" + ErrorMessage.GuestAdminMessage));

    // Act
    try {
      await client.addDatabaseUser();
    } catch (error) {
      // Assert
      chai.assert.include(error.message, ErrorMessage.GuestAdminError);
    }
  });

  it("initToken no provider error", async function () {
    // Arrange
    pluginContext.azureAccountProvider!.getIdentityCredentialAsync = async () => undefined;

    // Act
    try {
      await client.initToken();
    } catch (error) {
      // Assert
      const reason = ErrorMessage.IdentityCredentialUndefine(
        sqlPlugin.sqlImpl.config.identity,
        sqlPlugin.sqlImpl.config.databaseName
      );
      chai.assert.include(error.message, reason);
    }
  });

  it("initToken token error", async function () {
    // Arrange
    sinon.stub(SqlClient.prototype, "doQuery").rejects(new Error("test error"));

    // Act
    try {
      await client.initToken();
    } catch (error) {
      // Assert
      chai.assert.include(error.message, ErrorMessage.GetDetail);
    }
  });

  it("initToken error with domain code", async function () {
    // Arrange
    sinon
      .stub(msRestNodeAuth.ApplicationTokenCredentials.prototype, "getToken")
      .rejects(new Error("test error" + ErrorMessage.DomainCode));

    // Act
    try {
      await client.initToken();
    } catch (error) {
      // Assert
      chai.assert.include(error.message, ErrorMessage.DomainError);
    }
  });
});
