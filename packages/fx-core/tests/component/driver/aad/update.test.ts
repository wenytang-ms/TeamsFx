// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import mockedEnv, { RestoreFn } from "mocked-env";
import chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { UpdateAadAppDriver } from "../../../../src/component/driver/aad/update";
import {
  MockedLogProvider,
  MockedM365Provider,
  MockedTelemetryReporter,
  MockedUserInteraction,
} from "../../../plugins/solution/util";
import { AadAppClient } from "../../../../src/component/driver/aad/utility/aadAppClient";
import path from "path";
import * as fs from "fs-extra";
import { MissingFieldInManifestUserError } from "../../../../src/component/driver/aad/error/invalidFieldInManifestError";
import {
  UnhandledSystemError,
  UnhandledUserError,
} from "../../../../src/component/driver/aad/error/unhandledError";
import { InvalidParameterUserError } from "../../../../src/component/driver/aad/error/invalidParameterUserError";
import { cwd } from "process";
import { MissingEnvUserError } from "../../../../src/component/driver/aad/error/missingEnvError";

chai.use(chaiAsPromised);
const expect = chai.expect;

const outputKeys = {
  AAD_APP_ACCESS_AS_USER_PERMISSION_ID: "AAD_APP_ACCESS_AS_USER_PERMISSION_ID",
};

const testAssetsRoot = "./tests/component/driver/aad/testAssets";
const outputRoot = path.join(testAssetsRoot, "output");

describe("aadAppUpdate", async () => {
  const expectedObjectId = "00000000-0000-0000-0000-000000000000";
  const expectedClientId = "00000000-0000-0000-0000-111111111111";
  const expectedPermissionId = "00000000-0000-0000-0000-222222222222";
  const updateAadAppDriver = new UpdateAadAppDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    projectPath: cwd(),
    ui: new MockedUserInteraction(),
  };

  let envRestore: RestoreFn | undefined;

  afterEach(async () => {
    sinon.restore();
    if (envRestore) {
      envRestore();
      envRestore = undefined;
    }
    await fs.remove(outputRoot);
  });

  it("should throw error if argument property is missing", async () => {
    let args: any = {};

    let result = await updateAadAppDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(InvalidParameterUserError)
      .and.has.property(
        "message",
        "Following parameter is missing or invalid for aadApp/update action: manifestTemplatePath, outputFilePath."
      );

    args = {
      manifestTemplatePath: "./aad.manifest.json",
    };

    result = await updateAadAppDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(InvalidParameterUserError)
      .and.has.property(
        "message",
        "Following parameter is missing or invalid for aadApp/update action: outputFilePath."
      );

    args = {
      outputFilePath: "./build/aad.manifest.dev.json",
    };

    result = await updateAadAppDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(InvalidParameterUserError)
      .and.has.property(
        "message",
        "Following parameter is missing or invalid for aadApp/update action: manifestTemplatePath."
      );
  });

  it("should throw error if argument property is invalid", async () => {
    let args: any = {
      manifestTempaltePath: "",
      outputFilePath: "./build/aad.manifest.dev.json",
    };

    let result = await updateAadAppDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(InvalidParameterUserError)
      .and.has.property(
        "message",
        "Following parameter is missing or invalid for aadApp/update action: manifestTemplatePath."
      );

    args = {
      manifestTemplatePath: "./aad.manifest.json",
      outputFilePath: "",
    };

    result = await updateAadAppDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(InvalidParameterUserError)
      .and.has.property(
        "message",
        "Following parameter is missing or invalid for aadApp/update action: outputFilePath."
      );

    args = {
      manifestTemplatePath: true,
      outputFilePath: true,
    };

    result = await updateAadAppDriver.execute(args, mockedDriverContext);
    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(InvalidParameterUserError)
      .and.has.property(
        "message",
        "Following parameter is missing or invalid for aadApp/update action: manifestTemplatePath, outputFilePath."
      );
  });

  it("should success with valid manifest", async () => {
    sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    const outputPath = path.join(outputRoot, "manifest.output.json");
    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: outputPath,
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isOk()).to.be.true;
    expect(result.result._unsafeUnwrap().get(outputKeys.AAD_APP_ACCESS_AS_USER_PERMISSION_ID)).to.be
      .not.empty;
    expect(result.result._unsafeUnwrap().size).to.equal(1);
    expect(await fs.pathExists(path.join(outputPath))).to.be.true;
    const actualManifest = JSON.parse(await fs.readFile(outputPath, "utf8"));
    expect(actualManifest.id).to.equal(expectedObjectId);
    expect(actualManifest.appId).to.equal(expectedClientId);
    expect(actualManifest.requiredResourceAccess[0].resourceAppId).to.equal(
      "00000003-0000-0000-c000-000000000000"
    ); // Should convert Microsoft Graph to its id
    expect(actualManifest.requiredResourceAccess[0].resourceAccess[0].id).to.equal(
      "e1fe6dd8-ba31-4d61-89e7-88639da4683d"
    ); // Should convert User.Read to its id
    expect(actualManifest.oauth2Permissions[0].id).to.not.equal(
      "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
    ); // Should be replaced with an actual value
    expect(result.summaries.length).to.equal(1);
    console.log(result.summaries[0]);
    expect(result.summaries).includes(
      `Applied manifest ${args.manifestTemplatePath} to Azure Active Directory application with object id ${expectedObjectId}`
    );
  });

  it("should use absolute path in args directly", async () => {
    const outputPath = path.join(cwd(), outputRoot, "manifest.output.json");
    const manifestPath = path.join(cwd(), testAssetsRoot, "manifest.json");
    process.chdir("tests"); // change cwd for test
    try {
      sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
      envRestore = mockedEnv({
        AAD_APP_OBJECT_ID: expectedObjectId,
        AAD_APP_CLIENT_ID: expectedClientId,
      });

      const args = {
        manifestTemplatePath: manifestPath,
        outputFilePath: outputPath,
      };

      const result = await updateAadAppDriver.execute(args, mockedDriverContext);

      expect(result.result.isOk()).to.be.true;
    } finally {
      process.chdir(".."); // restore cwd
    }
  });

  it("should add project path to relative path in args", async () => {
    process.chdir("tests"); // change cwd for test
    try {
      sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
      envRestore = mockedEnv({
        AAD_APP_OBJECT_ID: expectedObjectId,
        AAD_APP_CLIENT_ID: expectedClientId,
      });

      const args = {
        manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
        outputFilePath: path.join(outputRoot, "manifest.output.json"),
      };

      const result = await updateAadAppDriver.run(args, mockedDriverContext);

      expect(result.isOk()).to.be.true;
    } finally {
      process.chdir(".."); // restore cwd
    }
  });

  it("should throw error if manifest does not contain id", async () => {
    sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
    envRestore = mockedEnv({
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    let args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };

    let result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).is.instanceOf(MissingEnvUserError).and.include({
      message:
        "Failed to generate AAD app manifest. Environment variable AAD_APP_OBJECT_ID is not set.", // The env does not have AAD_APP_OBJECT_ID so the id value is invalid
      source: "aadApp/update",
    });

    args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifestWithoutId.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };

    result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(MissingFieldInManifestUserError)
      .and.include({
        message: "Field id is missing or invalid in AAD app manifest.", // The manifest does not has an id property
        source: "aadApp/update",
      });
  });

  it("should only call MS Graph API once if manifest does not have preAuthorizedApplications", async () => {
    sinon
      .stub(AadAppClient.prototype, "updateAadApp")
      .onCall(0)
      .resolves()
      .onCall(1)
      .rejects("updateAadApp should not be called twice");

    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifestWithoutPreAuthorizedApp.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isOk()).to.be.true;
  });

  it("should call MS Graph API twice if manifest has preAuthorizedApplications", async () => {
    let requestCount = 0;
    sinon
      .stub(AadAppClient.prototype, "updateAadApp")
      .onCall(0)
      .callsFake(async (manifest) => {
        requestCount++;
        expect(manifest.preAuthorizedApplications.length).to.equal(0); // should have no preAuthorizedApplication in first request
      })
      .onCall(1)
      .callsFake(async (manifest) => {
        requestCount++;
        expect(manifest.preAuthorizedApplications.length).to.greaterThan(0); // should have preAuthorizedApplication in second request
      });

    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isOk()).to.be.true;
    expect(requestCount).to.equal(2); // should call MS Graph API twice
  });

  it("should not generate new permission id if the value already exists", async () => {
    sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
      AAD_APP_ACCESS_AS_USER_PERMISSION_ID: expectedPermissionId,
    });

    const outputPath = path.join(outputRoot, "manifest.output.json");
    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: outputPath,
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    const actualManifest = JSON.parse(await fs.readFile(outputPath, "utf8"));

    expect(result.result.isOk()).to.be.true;
    expect(
      result.result._unsafeUnwrap().get(outputKeys.AAD_APP_ACCESS_AS_USER_PERMISSION_ID)
    ).to.equal(expectedPermissionId);
    expect(result.result._unsafeUnwrap().size).to.equal(1);
    expect(actualManifest.oauth2Permissions[0].id).to.equal(expectedPermissionId);
  });

  it("should not generate new permission id if manifest does not need it", async () => {
    sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
      MY_PERMISSION_ID: expectedPermissionId,
    });

    const outputPath = path.join(outputRoot, "manifest.output.json");
    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifestWithNoPermissionId.json"),
      outputFilePath: outputPath,
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    const actualManifest = JSON.parse(await fs.readFile(outputPath, "utf8"));

    expect(result.result.isOk()).to.be.true;
    expect(result.result._unsafeUnwrap().size).to.equal(0);
    expect(actualManifest.oauth2Permissions[0].id).to.equal(expectedPermissionId);
  });

  it("should throw user error when AadAppClient failed with 4xx error", async () => {
    sinon.stub(AadAppClient.prototype, "updateAadApp").rejects({
      isAxiosError: true,
      response: {
        status: 400,
        data: {
          error: {
            code: "Request_BadRequest",
            message:
              "Invalid value specified for property 'displayName' of resource 'Application'.",
          },
        },
      },
    });
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(UnhandledUserError)
      .and.property("message")
      .contain("Unhandled error happened in aadApp/update action");
  });

  it("should throw system error when AadAppClient failed with non 4xx error", async () => {
    sinon.stub(AadAppClient.prototype, "updateAadApp").rejects({
      isAxiosError: true,
      response: {
        status: 500,
        data: {
          error: {
            code: "InternalServerError",
            message: "Internal server error",
          },
        },
      },
    });
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr())
      .is.instanceOf(UnhandledSystemError)
      .and.property("message")
      .contain("Unhandled error happened in aadApp/update action");
  });

  it("should send telemetries when success", async () => {
    const mockedTelemetryReporter = new MockedTelemetryReporter();
    let startTelemetry: any, endTelemetry: any;

    sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    sinon
      .stub(mockedTelemetryReporter, "sendTelemetryEvent")
      .onFirstCall()
      .callsFake((eventName, properties, measurements) => {
        startTelemetry = {
          eventName,
          properties,
          measurements,
        };
      })
      .onSecondCall()
      .callsFake((eventName, properties, measurements) => {
        endTelemetry = {
          eventName,
          properties,
          measurements,
        };
      });

    const outputPath = path.join(outputRoot, "manifest.output.json");
    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: outputPath,
    };
    const dirverContext: any = {
      m365TokenProvider: new MockedM365Provider(),
      logProvider: new MockedLogProvider(),
      telemetryReporter: mockedTelemetryReporter,
      projectPath: cwd(),
    };

    const result = await updateAadAppDriver.execute(args, dirverContext);

    expect(result.result.isOk()).to.be.true;
    expect(startTelemetry.eventName).to.equal("aadApp/update-start");
    expect(startTelemetry.properties.component).to.equal("aadApp/update");
    expect(endTelemetry.eventName).to.equal("aadApp/update");
    expect(endTelemetry.properties.component).to.equal("aadApp/update");
    expect(endTelemetry.properties.success).to.equal("yes");
  });

  it("should send error telemetries when fail", async () => {
    const mockedTelemetryReporter = new MockedTelemetryReporter();
    let startTelemetry: any, endTelemetry: any;

    sinon
      .stub(mockedTelemetryReporter, "sendTelemetryEvent")
      .onFirstCall()
      .callsFake((eventName, properties, measurements) => {
        startTelemetry = {
          eventName,
          properties,
          measurements,
        };
      });

    sinon
      .stub(mockedTelemetryReporter, "sendTelemetryErrorEvent")
      .onFirstCall()
      .callsFake((eventName, properties, measurements) => {
        endTelemetry = {
          eventName,
          properties,
          measurements,
        };
      });

    sinon.stub(AadAppClient.prototype, "updateAadApp").rejects({
      isAxiosError: true,
      response: {
        status: 500,
        data: {
          error: {
            code: "InternalServerError",
            message: "Internal server error",
          },
        },
      },
    });
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifest.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };
    const dirverContext: any = {
      m365TokenProvider: new MockedM365Provider(),
      logProvider: new MockedLogProvider(),
      telemetryReporter: mockedTelemetryReporter,
      projectPath: cwd(),
    };

    const result = await updateAadAppDriver.execute(args, dirverContext);

    expect(result.result.isOk()).to.be.false;
    expect(startTelemetry.eventName).to.equal("aadApp/update-start");
    expect(startTelemetry.properties.component).to.equal("aadApp/update");
    expect(endTelemetry.eventName).to.equal("aadApp/update");
    expect(endTelemetry.properties.component).to.equal("aadApp/update");
    expect(endTelemetry.properties.success).to.equal("no");
    expect(endTelemetry.properties["error-code"]).to.equal("aadApp/update.UnhandledError");
    expect(endTelemetry.properties["error-type"]).to.equal("system");
    expect(endTelemetry.properties["error-message"])
      .contain("Unhandled error happened in aadApp/update action")
      .and.contain("Internal server error");
  });

  it("should throw error when missing required environment variable in manifest", async () => {
    sinon.stub(AadAppClient.prototype, "updateAadApp").resolves();
    envRestore = mockedEnv({
      AAD_APP_OBJECT_ID: expectedObjectId,
      AAD_APP_CLIENT_ID: expectedClientId,
    });

    const args = {
      manifestTemplatePath: path.join(testAssetsRoot, "manifestWithMissingEnv.json"),
      outputFilePath: path.join(outputRoot, "manifest.output.json"),
    };

    const result = await updateAadAppDriver.execute(args, mockedDriverContext);

    expect(result.result.isErr()).to.be.true;
    expect(result.result._unsafeUnwrapErr()).is.instanceOf(MissingEnvUserError).and.include({
      message:
        "Failed to generate AAD app manifest. Environment variable AAD_APP_NAME, APPLICATION_NAME is not set.",
      source: "aadApp/update",
    });
  });
});
