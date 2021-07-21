// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

export enum TelemetryEvent {
  //TODO: define CLI telemetry event
  AccountLoginStart = "login-start",
  AccountLogin = "login",

  CreateProjectStart = "create-project-start",
  CreateProject = "create-project",

  DownloadSampleStart = "download-sample-start",
  DownloadSample = "download-sample",

  UpdateProjectStart = "add-resource-start",
  UpdateProject = "add-resource",

  AddCapStart = "add-capability-start",
  AddCap = "add-capability",

  ValidateManifestStart = "validate-manifest-start",
  ValidateManifest = "validate-manifest",

  BuildStart = "build-start",
  Build = "build",

  ProvisionStart = "provision-start",
  Provision = "provision",

  DeployStart = "deploy-start",
  Deploy = "deploy",

  PublishStart = "publish-start",
  Publish = "publish",

  InitStart = "init-start",
  Init = "init",

  PreviewStart = "preview-start",
  Preview = "preview",
  PreviewNpmInstallStart = "preview-npm-install-start",
  PreviewNpmInstall = "preview-npm-install",
  PreviewServiceStart = "preview-service-start",
  PreviewService = "preview-service",
  PreviewSideloadingStart = "preview-sideloading-start",
  PreviewSideloading = "preview-sideloading",

  ConfigGet = "config-get",
  ConfigSet = "config-set",
}

export enum TelemetryProperty {
  Component = "component",
  AppId = "appid",
  UserId = "hashed-userid",
  AccountType = "account-type",
  Success = "success",
  ErrorType = "error-type",
  ErrorCode = "error-code",
  ErrorMessage = "error-message",
  SampleName = "sample-name",
  Capabilities = "capabilities",
  Resources = "resources",
  Internal = "internal",
  InternalAlias = "internal-alias",
  PreviewAppId = "preview-appid",
  PreviewType = "preview-type",
  PreviewBrowser = "preview-browser",
  PreviewNpmInstallName = "preview-npm-install-name",
  PreviewNpmInstallExitCode = "preview-npm-install-exit-code",
  PreviewNpmInstallNodeVersion = "preview-npm-install-node-version",
  PreviewNpmInstallNpmVersion = "preview-npm-install-npm-version",
  PreviewNpmInstallErrorMessage = "preview-npm-install-error-message",
  PreviewServiceName = "preview-service-name",
  PreviewOSArch = "preview-os-arch",
  PreviewOSRelease = "preview-os-release",
}

export enum TelemetrySuccess {
  Yes = "yes",
  No = "no",
}

export enum TelemetryErrorType {
  UserError = "user",
  SystemError = "system",
}

export enum TelemetryAccountType {
  Azure = "azure",
  M365 = "m365",
}

export const TelemetryComponentType = "cli";
