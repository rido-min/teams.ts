// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { BotHttpClient, type TokenProvider } from './bot-http-client.js'
import type { CoreActivity } from '../schema/core-activity.js'
import { getLogger } from '../logging/logger.js'

const TOKEN_SERVICE_URL = 'https://token.botframework.com'

/** The OAuth token status for a user on a specific connection. */
export interface TokenStatus {
  /** Channel ID the token belongs to. */
  channelId: string;
  /** Name of the OAuth connection. */
  connectionName: string;
  /** Whether the user currently has a cached token. */
  hasToken: boolean;
  /** Display name of the OAuth service provider. */
  serviceProviderDisplayName?: string;
}

/** An OAuth token for a user. */
export interface UserToken {
  /** Channel ID the token belongs to. */
  channelId: string;
  /** Name of the OAuth connection. */
  connectionName: string;
  /** The access token string. */
  token: string;
  /** ISO-8601 expiry time of the token. */
  expiration?: string;
  /** Additional token properties. */
  properties?: Record<string, unknown>;
}

/** Resource returned by the Bot Framework sign-in endpoint. */
export interface SignInResource {
  /** URL to send the user to for sign-in. */
  signInLink?: string;
  /** Token exchange resource for SSO flows. */
  tokenExchangeResource?: TokenExchangeResource;
  /** Token post resource for direct token posting. */
  tokenPostResource?: TokenPostResource;
}

/** Metadata for an SSO token exchange. */
export interface TokenExchangeResource {
  /** Unique identifier for this exchange resource. */
  id: string;
  /** URI used for token exchange. */
  uri?: string;
  /** Identity provider ID. */
  providerId?: string;
}

/** SAS URL for direct token posting. */
export interface TokenPostResource {
  /** Shared Access Signature URL. */
  sasUrl?: string;
}

/** Request body for exchanging an SSO token. */
export interface TokenExchangeRequest {
  /** URI of the resource to exchange a token for. */
  uri?: string;
  /** The token to exchange. */
  token?: string;
}

/** A set of AAD resource URLs to fetch tokens for. */
export interface AadResourceUrls {
  /** List of resource URIs. */
  resourceUrls: string[];
}

/**
 * Client for the Bot Framework token service.
 *
 * Provides methods for OAuth sign-in flows, token retrieval, SSO token
 * exchange, user sign-out, and AAD token acquisition.
 */
export class UserTokenClient {
  private readonly http: BotHttpClient

  /**
   * @param getToken - Optional token provider for authenticating outgoing requests.
   * @param logger - Optional logger instance.
   */
  constructor (getToken?: TokenProvider) {
    this.http = new BotHttpClient(getToken)
  }

  /**
   * Retrieve the OAuth token status for a user across one or more connections.
   *
   * @param userId - ID of the user.
   * @param channelId - Channel the user is interacting on.
   * @param connectionName - Optional connection name to filter results.
   * @returns Array of token status records; empty if none found.
   */
  async getTokenStatusAsync (
    userId: string,
    channelId: string,
    connectionName?: string
  ): Promise<TokenStatus[]> {
    getLogger().info('Calling API endpoint: GetTokenStatus')
    const params: Record<string, string | undefined> = {
      userId,
      channelId,
      include: connectionName,
    }
    const result = await this.http.get<TokenStatus[]>(
      TOKEN_SERVICE_URL,
      '/api/usertoken/GetTokenStatus',
      params,
      { operationDescription: 'get token status' }
    )
    return result ?? []
  }

  /**
   * Retrieve a stored OAuth token for a user.
   *
   * @param userId - ID of the user.
   * @param channelId - Channel the user is on.
   * @param connectionName - Name of the OAuth connection.
   * @param code - Magic code from the sign-in flow (optional).
   * @returns The user token, or `undefined` if no token is stored.
   */
  async getTokenAsync (
    userId: string,
    channelId: string,
    connectionName: string,
    code?: string
  ): Promise<UserToken | undefined> {
    getLogger().info('Calling API endpoint: GetToken')
    const params: Record<string, string | undefined> = {
      userId,
      channelId,
      connectionName,
      code,
    }
    return this.http.get<UserToken>(
      TOKEN_SERVICE_URL,
      '/api/usertoken/GetToken',
      params,
      { operationDescription: 'get token', returnNullOnNotFound: true }
    )
  }

  /**
   * Retrieve the sign-in resource (URL and SSO exchange info) for a connection.
   *
   * @param connectionName - Name of the OAuth connection.
   * @param activity - The current incoming activity (used to build the state parameter).
   * @param finalRedirect - Optional URL to redirect to after sign-in completes.
   * @returns Sign-in resource, or `undefined` if unavailable.
   */
  async getSignInResourceAsync (
    connectionName: string,
    activity: CoreActivity,
    finalRedirect?: string
  ): Promise<SignInResource | undefined> {
    getLogger().info('Calling API endpoint: GetSignInResource')
    const params: Record<string, string | undefined> = {
      state: buildStateParam(connectionName, activity),
      finalRedirect,
    }
    return this.http.get<SignInResource>(
      TOKEN_SERVICE_URL,
      '/api/botsignin/GetSignInResource',
      params,
      { operationDescription: 'get sign-in resource' }
    )
  }

  /**
   * Exchange an SSO token for a user token on a given connection.
   *
   * @param userId - ID of the user.
   * @param channelId - Channel the user is on.
   * @param connectionName - Name of the OAuth connection.
   * @param request - Token exchange request (URI or raw token).
   * @returns The exchanged user token, or `undefined` if exchange fails.
   */
  async exchangeTokenAsync (
    userId: string,
    channelId: string,
    connectionName: string,
    request: TokenExchangeRequest
  ): Promise<UserToken | undefined> {
    getLogger().info('Calling API endpoint: ExchangeToken')
    const params: Record<string, string | undefined> = {
      userId,
      channelId,
      connectionName,
    }
    const url = buildTokenUrl(TOKEN_SERVICE_URL, '/api/usertoken/ExchangeToken', params)
    return this.http.send<UserToken>('POST', url, request, {
      operationDescription: 'exchange token',
    })
  }

  /**
   * Sign out a user from an OAuth connection.
   *
   * @param userId - ID of the user to sign out.
   * @param channelId - Channel the user is on.
   * @param connectionName - Optional connection name to sign out from. Omit to sign out of all connections.
   */
  async signOutUserAsync (
    userId: string,
    channelId: string,
    connectionName?: string
  ): Promise<void> {
    getLogger().info('Calling API endpoint: SignOut')
    const params: Record<string, string | undefined> = {
      userId,
      channelId,
      connectionName,
    }
    await this.http.delete(
      TOKEN_SERVICE_URL,
      '/api/usertoken/SignOut',
      params,
      { operationDescription: 'sign out user' }
    )
  }

  /**
   * Retrieve AAD tokens for multiple resources for a user.
   *
   * @param userId - ID of the user.
   * @param channelId - Channel the user is on.
   * @param connectionName - Name of the AAD OAuth connection.
   * @param resourceUrls - List of AAD resource URIs to fetch tokens for.
   * @returns Map from resource URL to user token; empty object if none returned.
   */
  async getAadTokensAsync (
    userId: string,
    channelId: string,
    connectionName: string,
    resourceUrls: AadResourceUrls
  ): Promise<Record<string, UserToken>> {
    getLogger().info('Calling API endpoint: GetAadTokens')
    const params: Record<string, string | undefined> = {
      userId,
      channelId,
      connectionName,
    }
    const url = buildTokenUrl(TOKEN_SERVICE_URL, '/api/usertoken/GetAadTokens', params)
    const result = await this.http.send<Record<string, UserToken>>(
      'POST',
      url,
      resourceUrls,
      { operationDescription: 'get AAD tokens' }
    )
    return result ?? {}
  }
}

function buildStateParam (connectionName: string, activity: CoreActivity): string {
  const state = {
    ConnectionName: connectionName,
    Conversation: {
      ActivityId: activity.id,
      Bot: activity.recipient,
      ChannelId: activity.channelId,
      Conversation: activity.conversation,
      ServiceUrl: activity.serviceUrl,
    },
    RelatesTo: null,
    MSAppId: activity.recipient?.id,
  }
  return Buffer.from(JSON.stringify(state)).toString('base64')
}

function buildTokenUrl (
  base: string,
  path: string,
  params: Record<string, string | undefined>
): string {
  const query = Object.entries(params)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v!)}`)
    .join('&')
  const url = `${base}${path}`
  return query ? `${url}?${query}` : url
}
