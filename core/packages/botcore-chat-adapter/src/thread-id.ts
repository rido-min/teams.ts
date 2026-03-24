// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Platform-specific data encoded in a Teams thread ID.
 * Both fields are required to make Bot Framework API calls.
 */
export interface TeamsThreadId {
  /** Bot Framework service URL (e.g. `https://smba.trafficmanager.net/amer/`) */
  serviceUrl: string
  /** Teams conversation ID (e.g. `19:abc123@thread.tacv2`) */
  conversationId: string
}

/**
 * Encode a Teams thread ID from platform-specific data.
 * Format: `teams:{base64url(serviceUrl)}:{base64url(conversationId)}`
 */
export function encodeThreadId (data: TeamsThreadId): string {
  const svc = Buffer.from(data.serviceUrl).toString('base64url')
  const conv = Buffer.from(data.conversationId).toString('base64url')
  return `teams:${svc}:${conv}`
}

/**
 * Decode a Teams thread ID back to its platform-specific components.
 * Throws if the thread ID is not a valid Teams thread ID.
 */
export function decodeThreadId (threadId: string): TeamsThreadId {
  const idx1 = threadId.indexOf(':')
  const idx2 = threadId.indexOf(':', idx1 + 1)
  if (idx1 === -1 || idx2 === -1 || threadId.slice(0, idx1) !== 'teams') {
    throw new Error(`Invalid Teams thread ID: ${threadId}`)
  }
  const svc = Buffer.from(threadId.slice(idx1 + 1, idx2), 'base64url').toString('utf8')
  const conv = Buffer.from(threadId.slice(idx2 + 1), 'base64url').toString('utf8')
  return { serviceUrl: svc, conversationId: conv }
}
