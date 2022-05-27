// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken, GetTokenOptions, TokenCredential } from "@azure/identity";
import { UserInfo } from "../models/userinfo";
import { AuthenticationConfiguration } from "../models/configuration";
import { formatString } from "../util/utils";
import { ErrorWithCode, ErrorCode, ErrorMessage } from "../core/errors";

/**
 * Represent on-behalf-of flow to get user identity, and it is designed to be used in Azure Function or Bot scenarios.
 *
 * @remarks
 * Can only be used in server side.
 */
export class OnBehalfOfUserCredential implements TokenCredential {
  /**
   * Constructor of OnBehalfOfUserCredential
   *
   * @remarks
   * Can Only works in in server side.
   */
  constructor(ssoToken: string, config: AuthenticationConfiguration) {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "OnBehalfOfUserCredential"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Get access token from credential.
   * @remarks
   * Can only be used in server side.
   */
  async getToken(
    scopes: string | string[],
    options?: GetTokenOptions
  ): Promise<AccessToken | null> {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "OnBehalfOfUserCredential"),
      ErrorCode.RuntimeNotSupported
    );
  }

  /**
   * Get basic user info from SSO token.
   * @remarks
   * Can only be used in server side.
   */
  public getUserInfo(): Promise<UserInfo> {
    throw new ErrorWithCode(
      formatString(ErrorMessage.BrowserRuntimeNotSupported, "OnBehalfOfUserCredential"),
      ErrorCode.RuntimeNotSupported
    );
  }
}
