// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { BotHttpClient, type BotRequestOptions, type TokenProvider } from './bot-http-client.js'
import { getLogger } from '../logging/logger.js'
import type {
  ChannelAccount,
  ConversationAccount,
  ConversationParameters,
  ConversationResourceResponse,
  ConversationsResult,
  CoreActivity,
  PagedMembersResult,
  ResourceResponse,
  Transcript,
} from '../schema/core-activity.js'

/** Binary attachment data for uploading to the Bot Framework attachment store. */
export interface AttachmentData {
  /** MIME type of the attachment. */
  type: string;
  /** Filename. */
  name: string;
  /** Base64-encoded original file content. */
  originalBase64: string;
  /** Base64-encoded thumbnail image (optional). */
  thumbnailBase64?: string;
}

/**
 * Client for the Bot Framework v3 Conversations REST API.
 *
 * Wraps all conversation-scoped endpoints: sending and managing activities,
 * listing and modifying members, creating conversations, and uploading attachments.
 */
export class ConversationClient {
  private readonly http: BotHttpClient

  /**
   * @param getToken - Optional token provider used to authenticate outgoing requests.
   * @param logger - Optional logger instance.
   */
  constructor (getToken?: TokenProvider) {
    this.http = new BotHttpClient(getToken)
  }

  /**
   * Send an activity to a conversation.
   *
   * @param serviceUrl - Bot Framework service URL (from the incoming activity).
   * @param conversationId - Target conversation ID.
   * @param activity - Activity payload to send.
   * @returns The resource response containing the new activity ID.
   */
  async sendActivityAsync (
    serviceUrl: string,
    conversationId: string,
    activity: Partial<CoreActivity>
  ): Promise<ResourceResponse | undefined> {
    const isTargeted = activity.recipient?.isTargeted
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/activities${isTargeted ? '?isTargetedActivity=true' : ''}`
    getLogger().trace('Sending activity to %s', url)
    return this.http.post<ResourceResponse>(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/activities`,
      activity,
      { operationDescription: 'send activity' },
      isTargeted ? { isTargetedActivity: 'true' } : undefined
    )
  }

  /**
   * Update an existing activity in a conversation.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param activityId - ID of the activity to update.
   * @param activity - Updated activity payload.
   * @returns The resource response.
   */
  async updateActivityAsync (
    serviceUrl: string,
    conversationId: string,
    activityId: string,
    activity: Partial<CoreActivity>
  ): Promise<ResourceResponse | undefined> {
    const isTargeted = activity.recipient?.isTargeted
    const endpoint = `/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}`
    const url = `${serviceUrl}${endpoint}${isTargeted ? '?isTargetedActivity=true' : ''}`
    getLogger().trace('Updating activity at %s', url)
    return this.http.put<ResourceResponse>(
      serviceUrl,
      endpoint,
      activity,
      { operationDescription: 'update activity' },
      isTargeted ? { isTargetedActivity: 'true' } : undefined
    )
  }

  /**
   * Update a Targeted Message activity.
   *
   * Equivalent to {@link updateActivityAsync} with `activity.recipient.isTargeted = true`,
   * but provided as a convenience method matching the C# `UpdateTargetedActivityAsync` API.
   * Always appends `?isTargetedActivity=true` to the request URL.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param activityId - ID of the activity to update.
   * @param activity - Updated activity payload.
   * @returns The resource response.
   */
  async updateTargetedActivityAsync (
    serviceUrl: string,
    conversationId: string,
    activityId: string,
    activity: Partial<CoreActivity>
  ): Promise<ResourceResponse | undefined> {
    const endpoint = `/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}`
    getLogger().trace('Updating targeted activity at %s%s', serviceUrl, endpoint)
    return this.http.put<ResourceResponse>(
      serviceUrl,
      endpoint,
      activity,
      { operationDescription: 'update targeted activity' },
      { isTargetedActivity: 'true' }
    )
  }

  /**
   * Delete an activity from a conversation.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param activityId - ID of the activity to delete.
   */
  async deleteActivityAsync (
    serviceUrl: string,
    conversationId: string,
    activityId: string
  ): Promise<void> {
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}`
    getLogger().trace('Deleting activity at %s', url)
    await this.http.delete(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}`,
      undefined,
      { operationDescription: 'delete activity' }
    )
  }

  /**
   * Delete a Targeted Message activity.
   *
   * Equivalent to {@link deleteActivityAsync} but always appends
   * `?isTargetedActivity=true` to the request URL, matching the C#
   * `DeleteTargetedActivityAsync` API.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param activityId - ID of the activity to delete.
   */
  async deleteTargetedActivityAsync (
    serviceUrl: string,
    conversationId: string,
    activityId: string
  ): Promise<void> {
    const endpoint = `/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}`
    getLogger().trace('Deleting targeted activity at %s%s', serviceUrl, endpoint)
    await this.http.delete(
      serviceUrl,
      endpoint,
      { isTargetedActivity: 'true' },
      { operationDescription: 'delete targeted activity' }
    )
  }

  /**
   * Retrieve all members of a conversation.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @returns Array of channel accounts; empty array if none found.
   */
  async getConversationMembersAsync (
    serviceUrl: string,
    conversationId: string
  ): Promise<ChannelAccount[]> {
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/members`
    getLogger().trace('Getting conversation members from %s', url)
    const result = await this.http.get<ChannelAccount[]>(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/members`,
      undefined,
      { operationDescription: 'get conversation members' }
    )
    return result ?? []
  }

  /**
   * Retrieve a single member of a conversation by user ID.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param memberId - User ID of the member to fetch.
   * @returns The channel account, or `undefined` if not found.
   */
  async getConversationMemberAsync (
    serviceUrl: string,
    conversationId: string,
    memberId: string
  ): Promise<ChannelAccount | undefined> {
    getLogger().trace('Getting conversation member from %s conversationId=%s memberId=%s', serviceUrl, conversationId, memberId)
    return this.http.get<ChannelAccount>(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/members/${memberId}`,
      undefined,
      { operationDescription: 'get conversation member', returnNullOnNotFound: true }
    )
  }

  /**
   * Retrieve a page of conversation members.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param pageSize - Maximum number of members to return per page.
   * @param continuationToken - Token from a previous response to fetch the next page.
   * @returns Paged result with members and optional continuation token.
   */
  async getConversationPagedMembersAsync (
    serviceUrl: string,
    conversationId: string,
    pageSize?: number,
    continuationToken?: string
  ): Promise<PagedMembersResult<ChannelAccount>> {
    getLogger().trace('Getting paged conversation members from %s conversationId=%s', serviceUrl, conversationId)
    const params: Record<string, string | undefined> = {
      pageSize: pageSize?.toString(),
      continuationToken,
    }
    const result = await this.http.get<PagedMembersResult<ChannelAccount>>(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/pagedmembers`,
      params,
      { operationDescription: 'get paged members' }
    )
    return result ?? { members: [] }
  }

  /**
   * Remove a member from a conversation.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param memberId - User ID of the member to remove.
   */
  async deleteConversationMemberAsync (
    serviceUrl: string,
    conversationId: string,
    memberId: string
  ): Promise<void> {
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/members/${memberId}`
    getLogger().trace('Deleting conversation member at %s', url)
    await this.http.delete(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/members/${memberId}`,
      undefined,
      { operationDescription: 'delete conversation member' }
    )
  }

  /**
   * Create a new proactive conversation.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param parameters - Conversation creation parameters.
   * @returns Resource response with the new conversation ID and service URL.
   */
  async createConversationAsync (
    serviceUrl: string,
    parameters: ConversationParameters
  ): Promise<ConversationResourceResponse | undefined> {
    getLogger().trace('Creating conversation at %s', serviceUrl)
    return this.http.post<ConversationResourceResponse>(
      serviceUrl,
      '/v3/conversations',
      parameters,
      { operationDescription: 'create conversation' }
    )
  }

  /**
   * List all conversations the bot is a member of.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param continuationToken - Token from a previous response to fetch the next page.
   * @returns Paged result with conversations and optional continuation token.
   */
  async getConversationsAsync (
    serviceUrl: string,
    continuationToken?: string
  ): Promise<ConversationsResult> {
    getLogger().trace('Getting conversations from %s', serviceUrl)
    const params: Record<string, string | undefined> = { continuationToken }
    const result = await this.http.get<ConversationsResult>(
      serviceUrl,
      '/v3/conversations',
      params,
      { operationDescription: 'get conversations' }
    )
    return result ?? { conversations: [] }
  }

  /**
   * Upload a transcript of past activities to a conversation.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param transcript - The transcript to upload.
   * @param options - Additional request options.
   * @returns Resource response.
   */
  async sendConversationHistoryAsync (
    serviceUrl: string,
    conversationId: string,
    transcript: Transcript,
    options?: BotRequestOptions
  ): Promise<ResourceResponse | undefined> {
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/activities/history`
    getLogger().trace('Sending conversation history to %s', url)
    return this.http.post<ResourceResponse>(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/activities/history`,
      transcript,
      { operationDescription: 'send conversation history', ...options }
    )
  }

  /**
   * Add a reaction to a message activity.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param activityId - ID of the message to react to.
   * @param reactionType - Reaction type string (e.g. `"like"`).
   */
  async addReactionAsync (
    serviceUrl: string,
    conversationId: string,
    activityId: string,
    reactionType: string
  ): Promise<void> {
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}/reactions`
    getLogger().trace('Adding reaction at %s', url)
    await this.http.post(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}/reactions`,
      { type: reactionType },
      { operationDescription: 'add reaction' }
    )
  }

  /**
   * Remove a reaction from a message activity.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @param activityId - ID of the message to remove the reaction from.
   * @param reactionType - Reaction type string to remove.
   */
  async deleteReactionAsync (
    serviceUrl: string,
    conversationId: string,
    activityId: string,
    reactionType: string
  ): Promise<void> {
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}/reactions/${reactionType}`
    getLogger().trace('Deleting reaction at %s', url)
    await this.http.delete(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/activities/${activityId}/reactions/${reactionType}`,
      undefined,
      { operationDescription: 'delete reaction' }
    )
  }

  /**
   * Upload a binary attachment to the Bot Framework attachment store.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID to scope the attachment.
   * @param attachment - Attachment data including base64 content.
   * @returns Resource response with the attachment ID.
   */
  async uploadAttachmentAsync (
    serviceUrl: string,
    conversationId: string,
    attachment: AttachmentData
  ): Promise<ResourceResponse | undefined> {
    const url = `${serviceUrl}/v3/conversations/${encodeConversationId(conversationId)}/attachments`
    getLogger().trace('Uploading attachment to %s', url)
    return this.http.post<ResourceResponse>(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}/attachments`,
      attachment,
      { operationDescription: 'upload attachment' }
    )
  }

  /**
   * Retrieve the conversation account details.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Conversation ID.
   * @returns The conversation account, or `undefined` if not found.
   */
  async getConversationAccountAsync (
    serviceUrl: string,
    conversationId: string
  ): Promise<ConversationAccount | undefined> {
    getLogger().trace('Getting conversation account from %s conversationId=%s', serviceUrl, conversationId)
    return this.http.get<ConversationAccount>(
      serviceUrl,
      `/v3/conversations/${encodeConversationId(conversationId)}`,
      undefined,
      { operationDescription: 'get conversation', returnNullOnNotFound: true }
    )
  }
}

// The 'agents' channel uses a long ID that must be truncated at the first semicolon
function encodeConversationId (conversationId: string): string {
  const truncated = conversationId.split(';')[0]
  if (truncated !== conversationId) {
    getLogger().info("Truncating conversation ID for 'agents' channel to comply with length restrictions")
  }
  return encodeURIComponent(truncated)
}
