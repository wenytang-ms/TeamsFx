// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { IConfigurableTab, IStaticTab, IMessagingExtensionCommand } from "@microsoft/teamsfx-api";

export interface IAppDefinition {
  teamsAppId?: string;
  tenantId?: string;
  ownerAadId?: string;
  userList?: IUserList[];
  environments?: any[];
  createdAt?: Date;
  updatedAt?: Date;
  appId?: string;
  appName: string;
  appStudioVersion?: string;
  version?: string;
  packageName?: string;
  shortName?: string;
  longName?: string;
  developerName?: string;
  websiteUrl?: string;
  privacyUrl?: string;
  termsOfUseUrl?: string;
  mpnId?: string;
  shortDescription?: string;
  longDescription?: string;
  colorIcon?: string;
  outlineIcon?: string;
  accentColor?: string;
  configurableTabs?: IConfigurableTab[];
  staticTabs?: IStaticTab[];
  bots?: IAppDefinitionBot[];
  connectors?: any[];
  messagingExtensions?: IMessagingExtension[];
  validDomains?: string[];
  appStudioChecklistChecked?: any[];
  webApplicationInfoId?: string;
  webApplicationInfoResource?: string;
  devicePermissions?: any[];
  applicationPermissions?: any[];
  showLoadingIndicator?: boolean;
  isFullScreen?: boolean;
  hasPreviewFeature?: boolean;
  localizationInfo?: ILocalizationInfo;
}

export interface IUserList {
  tenantId: string;
  aadId: string;
  displayName: string;
  userPrincipalName: string;
  isOwner: boolean;
}

export interface IAppDefinitionBot {
  objectId?: string;
  botId: string;
  needsChannelSelector?: boolean;
  isNotificationOnly: boolean;
  supportsFiles: boolean;
  isAudioCallingBot?: boolean;
  isVideoCallingBot?: boolean;
  scopes: string[];
  teamCommands?: ITeamCommand[];
  personalCommands?: IPersonalCommand[];
  groupChatCommands?: IGroupChatCommand[];
}

export interface ITeamCommand {
  title: string;
  description: string;
}

export interface IPersonalCommand {
  title: string;
  description: string;
}

export interface IGroupChatCommand {
  title: string;
  description: string;
}

export interface ILocalizationInfo {
  defaultLanguageTag?: any;
  languages: any[];
}

export interface IMessagingExtension {
  objectId?: string;
  botId: string;
  canUpdateConfiguration: boolean;
  commands: IMessagingExtensionCommand[];
  messageHandlers: {
    type: "link";
    value: {
      /**
       * A list of domains that the link message handler can register for, and when they are matched the app will be invoked
       */
      domains?: string[];
    };
  }[];
}
