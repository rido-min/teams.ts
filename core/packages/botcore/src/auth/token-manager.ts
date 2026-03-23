// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  ConfidentialClientApplication,
  ManagedIdentityApplication,
  type AuthenticationResult,
  type LogLevel as MSALLogLevel,
} from '@azure/msal-node'
import { getLogger } from '../logging/logger.js'

const BOT_TOKEN_SCOPE = 'https://api.botframework.com/.default'
const BOT_TOKEN_TENANT = 'botframework.com'
const AUTHORITY_BASE = 'https://login.microsoftonline.com'

export type TokenManagerOptions = {
  /** Application (client) ID. Falls back to CLIENT_ID env var. */
  readonly clientId?: string
  /** Client secret. Falls back to CLIENT_SECRET env var. */
  readonly clientSecret?: string
  /** Tenant ID. Falls back to TENANT_ID env var. */
  readonly tenantId?: string
  /**
   * Custom token factory. When provided, called instead of MSAL for every
   * token acquisition.
   */
  readonly token?: (scope: string, tenantId: string) => Promise<string>
  /**
   * Managed identity client ID for federated identity credentials.
   * Use `"system"` for system-assigned managed identity.
   * Falls back to MANAGED_IDENTITY_CLIENT_ID env var.
   */
  managedIdentityClientId?: 'system' | (string & Record<never, never>)
}

type ResolvedOptions = Required<Omit<TokenManagerOptions, 'token'>> & {
  token?: TokenManagerOptions['token']
}

/**
 * Manages Bot Framework access token acquisition via MSAL.
 *
 * Supports four authentication flows, selected automatically based on the
 * options provided:
 *
 * 1. **Client credentials** — `clientId` + `clientSecret`
 * 2. **Custom token factory** — `token` callback
 * 3. **User managed identity** — `clientId` only (no secret)
 * 4. **Federated identity** — `clientId` + `managedIdentityClientId` (different IDs)
 *
 * Confidential client instances are cached per `clientId:tenantId` pair.
 */
export class TokenManager {
  private readonly opts: ResolvedOptions
  private confidentialClients: Record<string, ConfidentialClientApplication> = {}
  private managedIdentityClient: ManagedIdentityApplication | null = null

  constructor (options: TokenManagerOptions = {}) {
    this.opts = {
      clientId: options.clientId ?? process.env['CLIENT_ID'] ?? '',
      clientSecret: options.clientSecret ?? process.env['CLIENT_SECRET'] ?? '',
      tenantId: options.tenantId ?? process.env['TENANT_ID'] ?? '',
      token: options.token,
      managedIdentityClientId:
        options.managedIdentityClientId ??
        (process.env['MANAGED_IDENTITY_CLIENT_ID'] as ResolvedOptions['managedIdentityClientId']) ??
        '',
    }
  }

  /** Acquire a Bot Framework access token. */
  async getBotToken (): Promise<string | null> {
    const tenantId = this.opts.tenantId || BOT_TOKEN_TENANT
    return this.getToken(BOT_TOKEN_SCOPE, tenantId)
  }

  private async getToken (scope: string, tenantId: string): Promise<string | null> {
    const { clientId, clientSecret, token, managedIdentityClientId } = this.opts

    if (!clientId) {
      getLogger().warn('Running without auth — no clientId configured')
      return null
    }

    // Custom token factory
    if (token) {
      getLogger().debug('Acquiring token via custom factory scope=%s tenantId=%s', scope, tenantId)
      return token(scope, tenantId)
    }

    // Client secret → confidential client credentials
    if (clientSecret) {
      getLogger().debug('Configuring authentication with client secret clientId=%s tenantId=%s', clientId, tenantId)
      return this.getTokenWithClientCredentials(clientId, clientSecret, scope, tenantId)
    }

    // No secret → managed / federated identity
    const hasFederated =
      managedIdentityClientId &&
      managedIdentityClientId.toLowerCase() !== clientId.toLowerCase()

    if (hasFederated) {
      const identityType = managedIdentityClientId === 'system' ? 'System-Assigned' : 'User-Assigned'
      getLogger().debug('Configuring authentication with Federated Identity Credential (Managed Identity) with %s Managed Identity clientId=%s', identityType, clientId)
      return this.getTokenWithFederatedCredentials(
        clientId,
        managedIdentityClientId!,
        scope,
        tenantId
      )
    }

    getLogger().debug('Configuring authentication with User-Assigned Managed Identity clientId=%s', clientId)
    return this.getTokenWithManagedIdentity(clientId, scope)
  }

  private async getTokenWithClientCredentials (
    clientId: string,
    clientSecret: string,
    scope: string,
    tenantId: string
  ): Promise<string | null> {
    const client = this.getConfidentialClient(clientId, clientSecret, tenantId)
    getLogger().debug('Acquiring app-only token for scope: %s', scope)
    const result = await client.acquireTokenByClientCredential({ scopes: [scope] })
    getLogger().debug('Token acquired expiresOn=%s', result?.expiresOn)
    return this.unwrap(result)
  }

  private async getTokenWithManagedIdentity (
    clientId: string,
    scope: string
  ): Promise<string | null> {
    const resource = stripDefault(scope)
    const client = this.getOrCreateManagedIdentityClient({ clientId })
    getLogger().debug('Acquiring app-only token for scope: %s', scope)
    const result = await client.acquireToken({ resource })
    getLogger().debug('Token acquired expiresOn=%s', result?.expiresOn)
    return this.unwrap(result)
  }

  private async getTokenWithFederatedCredentials (
    clientId: string,
    managedIdentityClientId: string,
    scope: string,
    tenantId: string
  ): Promise<string | null> {
    const miClient = this.getOrCreateManagedIdentityClient(
      managedIdentityClientId === 'system'
        ? { system: true }
        : { userAssignedClientId: managedIdentityClientId }
    )

    getLogger().debug('Acquiring agentic token for AgenticAppId %s', clientId)
    const miToken = await miClient.acquireToken({
      resource: 'api://AzureADTokenExchange',
    })
    getLogger().debug('Agentic token acquired')

    const confidentialClient = new ConfidentialClientApplication({
      auth: {
        clientId,
        clientAssertion: miToken.accessToken,
        authority: `${AUTHORITY_BASE}/${tenantId}`,
      },
      system: { loggerOptions: this.msalLoggerOptions() },
    })

    getLogger().debug('Acquiring app-only token for scope: %s', scope)
    const result = await confidentialClient.acquireTokenByClientCredential({
      scopes: [scope],
    })
    getLogger().debug('Token acquired expiresOn=%s', result?.expiresOn)
    return this.unwrap(result)
  }

  private getConfidentialClient (
    clientId: string,
    clientSecret: string,
    tenantId: string
  ): ConfidentialClientApplication {
    const key = `${clientId}:${tenantId}`
    if (!this.confidentialClients[key]) {
      this.confidentialClients[key] = new ConfidentialClientApplication({
        auth: {
          clientId,
          clientSecret,
          authority: `${AUTHORITY_BASE}/${tenantId}`,
        },
        system: { loggerOptions: this.msalLoggerOptions() },
      })
    }
    return this.confidentialClients[key]
  }

  private getOrCreateManagedIdentityClient (
    identity:
      | { clientId: string }
      | { userAssignedClientId: string }
      | { system: true }
  ): ManagedIdentityApplication {
    if (!this.managedIdentityClient) {
      const params =
        'system' in identity
          ? undefined
          : 'clientId' in identity
            ? { userAssignedClientId: identity.clientId }
            : { userAssignedClientId: identity.userAssignedClientId }

      this.managedIdentityClient = new ManagedIdentityApplication({
        managedIdentityIdParams: params,
        system: { loggerOptions: this.msalLoggerOptions() },
      })
    }
    return this.managedIdentityClient
  }

  private unwrap (result: AuthenticationResult | null): string | null {
    if (!result) throw new Error('MSAL returned no token result')
    return result.accessToken
  }

  private msalLoggerOptions (): { logLevel: MSALLogLevel; loggerCallback: () => void; piiLoggingEnabled: boolean } {
    return {
      logLevel: 3 /* Warning */ as MSALLogLevel,
      loggerCallback: () => {},
      piiLoggingEnabled: false,
    }
  }
}

function stripDefault (scope: string): string {
  return scope.replace('/.default', '')
}
