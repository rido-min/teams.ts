// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  ConfidentialClientApplication,
  ManagedIdentityApplication,
  type AuthenticationResult,
  type LogLevel as MSALLogLevel,
} from '@azure/msal-node';
import createDebug from 'debug';

const debug = createDebug('teams:botcore:msal');

const BOT_TOKEN_SCOPE = 'https://api.botframework.com/.default';
const BOT_TOKEN_TENANT = 'botframework.com';
const AUTHORITY_BASE = 'https://login.microsoftonline.com';

export type TokenManagerOptions = {
  /** Application (client) ID. Falls back to CLIENT_ID env var. */
  readonly clientId?: string;
  /** Client secret. Falls back to CLIENT_SECRET env var. */
  readonly clientSecret?: string;
  /** Tenant ID. Falls back to TENANT_ID env var. */
  readonly tenantId?: string;
  /**
   * Custom token factory. When provided, called instead of MSAL for every
   * token acquisition.
   */
  readonly token?: (scope: string, tenantId: string) => Promise<string>;
  /**
   * Managed identity client ID for federated identity credentials.
   * Use `"system"` for system-assigned managed identity.
   * Falls back to MANAGED_IDENTITY_CLIENT_ID env var.
   */
  managedIdentityClientId?: 'system' | (string & Record<never, never>);
};

type ResolvedOptions = Required<Omit<TokenManagerOptions, 'token'>> & {
  token?: TokenManagerOptions['token'];
};

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
  private readonly opts: ResolvedOptions;
  private confidentialClients: Record<string, ConfidentialClientApplication> = {};
  private managedIdentityClient: ManagedIdentityApplication | null = null;

  constructor(options: TokenManagerOptions = {}) {
    this.opts = {
      clientId: options.clientId ?? process.env['CLIENT_ID'] ?? '',
      clientSecret: options.clientSecret ?? process.env['CLIENT_SECRET'] ?? '',
      tenantId: options.tenantId ?? process.env['TENANT_ID'] ?? '',
      token: options.token,
      managedIdentityClientId:
        options.managedIdentityClientId ??
        (process.env['MANAGED_IDENTITY_CLIENT_ID'] as ResolvedOptions['managedIdentityClientId']) ??
        '',
    };
  }

  /** Acquire a Bot Framework access token. */
  async getBotToken(): Promise<string | null> {
    const tenantId = this.opts.tenantId || BOT_TOKEN_TENANT;
    return this.getToken(BOT_TOKEN_SCOPE, tenantId);
  }

  private async getToken(scope: string, tenantId: string): Promise<string | null> {
    const { clientId, clientSecret, token, managedIdentityClientId } = this.opts;

    if (!clientId) {
      debug('no clientId configured, skipping token acquisition');
      return null;
    }

    // Custom token factory
    if (token) {
      debug('acquiring token via custom factory scope=%s tenantId=%s', scope, tenantId);
      return token(scope, tenantId);
    }

    // Client secret → confidential client credentials
    if (clientSecret) {
      debug('acquiring token via client credentials clientId=%s scope=%s tenantId=%s', clientId, scope, tenantId);
      return this.getTokenWithClientCredentials(clientId, clientSecret, scope, tenantId);
    }

    // No secret → managed / federated identity
    const hasFederated =
      managedIdentityClientId &&
      managedIdentityClientId.toLowerCase() !== clientId.toLowerCase();

    if (hasFederated) {
      debug('acquiring token via federated identity clientId=%s managedIdentityClientId=%s scope=%s', clientId, managedIdentityClientId, scope);
      return this.getTokenWithFederatedCredentials(
        clientId,
        managedIdentityClientId!,
        scope,
        tenantId
      );
    }

    debug('acquiring token via user managed identity clientId=%s scope=%s', clientId, scope);
    return this.getTokenWithManagedIdentity(clientId, scope);
  }

  private async getTokenWithClientCredentials(
    clientId: string,
    clientSecret: string,
    scope: string,
    tenantId: string
  ): Promise<string | null> {
    const client = this.getConfidentialClient(clientId, clientSecret, tenantId);
    debug('MSAL acquireTokenByClientCredential scope=%s', scope);
    const result = await client.acquireTokenByClientCredential({ scopes: [scope] });
    debug('MSAL token acquired expiresOn=%s', result?.expiresOn);
    return this.unwrap(result);
  }

  private async getTokenWithManagedIdentity(
    clientId: string,
    scope: string
  ): Promise<string | null> {
    const resource = stripDefault(scope);
    const client = this.getOrCreateManagedIdentityClient({ clientId });
    debug('MSAL managed identity acquireToken resource=%s', resource);
    const result = await client.acquireToken({ resource });
    debug('MSAL managed identity token acquired expiresOn=%s', result?.expiresOn);
    return this.unwrap(result);
  }

  private async getTokenWithFederatedCredentials(
    clientId: string,
    managedIdentityClientId: string,
    scope: string,
    tenantId: string
  ): Promise<string | null> {
    const miClient = this.getOrCreateManagedIdentityClient(
      managedIdentityClientId === 'system'
        ? { system: true }
        : { userAssignedClientId: managedIdentityClientId }
    );

    debug('MSAL federated: acquiring MI assertion token');
    const miToken = await miClient.acquireToken({
      resource: 'api://AzureADTokenExchange',
    });
    debug('MSAL federated: MI assertion token acquired');

    const confidentialClient = new ConfidentialClientApplication({
      auth: {
        clientId,
        clientAssertion: miToken.accessToken,
        authority: `${AUTHORITY_BASE}/${tenantId}`,
      },
      system: { loggerOptions: this.msalLoggerOptions() },
    });

    debug('MSAL federated: exchanging assertion for scope=%s', scope);
    const result = await confidentialClient.acquireTokenByClientCredential({
      scopes: [scope],
    });
    debug('MSAL federated: token acquired expiresOn=%s', result?.expiresOn);
    return this.unwrap(result);
  }

  private getConfidentialClient(
    clientId: string,
    clientSecret: string,
    tenantId: string
  ): ConfidentialClientApplication {
    const key = `${clientId}:${tenantId}`;
    if (!this.confidentialClients[key]) {
      this.confidentialClients[key] = new ConfidentialClientApplication({
        auth: {
          clientId,
          clientSecret,
          authority: `${AUTHORITY_BASE}/${tenantId}`,
        },
        system: { loggerOptions: this.msalLoggerOptions() },
      });
    }
    return this.confidentialClients[key];
  }

  private getOrCreateManagedIdentityClient(
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
            : { userAssignedClientId: identity.userAssignedClientId };

      this.managedIdentityClient = new ManagedIdentityApplication({
        managedIdentityIdParams: params,
        system: { loggerOptions: this.msalLoggerOptions() },
      });
    }
    return this.managedIdentityClient;
  }

  private unwrap(result: AuthenticationResult | null): string | null {
    if (!result) throw new Error('MSAL returned no token result');
    return result.accessToken;
  }

  private msalLoggerOptions(): { logLevel: MSALLogLevel; loggerCallback: () => void; piiLoggingEnabled: boolean } {
    return {
      logLevel: 3 /* Warning */ as MSALLogLevel,
      loggerCallback: () => {},
      piiLoggingEnabled: false,
    };
  }
}

function stripDefault(scope: string): string {
  return scope.replace('/.default', '');
}
