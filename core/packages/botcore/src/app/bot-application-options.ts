// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { TokenManagerOptions } from '../auth/token-manager.js';

/**
 * Options for BotApplication. Credentials are resolved in this priority order:
 *
 * | clientId | clientSecret | managedIdentityClientId | Result                         |
 * |----------|--------------|-------------------------|--------------------------------|
 * | not set  | —            | —                       | No auth (dev/testing only)     |
 * | set      | set          | —                       | Client secret                  |
 * | set      | not set      | —                       | User managed identity          |
 * | set      | not set      | same as clientId        | User managed identity          |
 * | set      | not set      | different               | Federated identity (UMI)       |
 * | set      | not set      | "system"                | Federated identity (system MI) |
 *
 * All fields fall back to their corresponding environment variables
 * (`CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`, `MANAGED_IDENTITY_CLIENT_ID`).
 */
export interface BotApplicationOptions extends TokenManagerOptions {}
