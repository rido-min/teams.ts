// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
  type Adapter,
  type AdapterPostableMessage,
  type ChatInstance,
  type EmojiValue,
  type FetchOptions,
  type FetchResult,
  type FormattedContent,
  type MessageData,
  type RawMessage,
  type ThreadInfo,
  type WebhookOptions,
  Message,
  NotImplementedError,
  convertEmojiPlaceholders,
  getEmoji,
  isCardElement,
  parseMarkdown,
  stringifyMarkdown,
} from 'chat'
import type {
  Attachment as BotAttachment,
  BotApplicationOptions,
  ChannelAccount,
  CoreActivity,
} from '@microsoft/teams.botcore'
import {
  BotApplication,
  BotAuthError,
  validateBotToken,
} from '@microsoft/teams.botcore'
import { decodeThreadId, encodeThreadId, type TeamsThreadId } from './thread-id.js'
import { TeamsFormatConverter, teamsHtmlToMarkdown } from './format-converter.js'
import { cardToAdaptiveCard } from './cards.js'

export interface TeamsAdapterOptions extends BotApplicationOptions {
  /** Display name reported as the bot's user name. Defaults to `"bot"`. */
  userName?: string
}

/**
 * Chat SDK adapter for Microsoft Teams.
 *
 * Bridges the {@link https://chat-sdk.dev chat} SDK and the Bot Framework
 * by delegating webhook verification and API calls to
 * `@microsoft/teams.botcore`.
 *
 * @example
 * ```ts
 * import { Chat } from 'chat'
 * import { TeamsAdapter } from '@microsoft/teams.botcore-chat-adapter'
 *
 * const chat = new Chat({
 *   adapter: new TeamsAdapter({ clientId: '...', clientSecret: '...' }),
 * })
 *
 * chat.onMessage(async (thread, message) => {
 *   await thread.post(`You said: ${message.text}`)
 * })
 *
 * // In your HTTP handler (Hono / Express / etc.):
 * // return adapter.handleWebhook(request)
 * ```
 */
export class TeamsAdapter implements Adapter<TeamsThreadId, CoreActivity> {
  readonly name = 'teams'
  readonly userName: string

  private chat: ChatInstance | undefined
  private readonly app: BotApplication
  private readonly converter: TeamsFormatConverter

  constructor (options: TeamsAdapterOptions = {}) {
    this.app = new BotApplication(options)
    this.userName = options.userName ?? 'bot'
    this.converter = new TeamsFormatConverter()
  }

  async initialize (chat: ChatInstance): Promise<void> {
    this.chat = chat
  }

  encodeThreadId (data: TeamsThreadId): string {
    return encodeThreadId(data)
  }

  decodeThreadId (threadId: string): TeamsThreadId {
    return decodeThreadId(threadId)
  }

  /**
   * Returns the "channel" portion of the thread ID.
   * For Teams this is `teams:{base64url(serviceUrl)}`.
   */
  channelIdFromThreadId (threadId: string): string {
    const idx1 = threadId.indexOf(':')
    const idx2 = threadId.indexOf(':', idx1 + 1)
    return threadId.slice(0, idx2)
  }

  /**
   * Returns true for personal (1:1) chats.
   *
   * In Teams, all channel and group-chat conversation IDs begin with `"19:"`.
   * Personal (1:1) DMs use a different format (e.g. `"8:orgid:..."`) and never
   * start with `"19:"`, so a single prefix check is sufficient.
   */
  isDM (threadId: string): boolean {
    const { conversationId } = decodeThreadId(threadId)
    return !conversationId.startsWith('19:')
  }

  /**
   * Handle an incoming Bot Framework webhook request.
   *
   * Validates the Bearer token, dispatches message/reaction/action activities
   * to the Chat SDK, and returns a `200 {}` response immediately (fire-and-forget
   * via `waitUntil`). Invoke activities return `invokeResponse` as required by
   * the Bot Framework protocol.
   */
  async handleWebhook (request: Request, options?: WebhookOptions): Promise<Response> {
    const authHeader = request.headers.get('authorization') ?? undefined
    try {
      await validateBotToken(authHeader, this.app.options.clientId)
    } catch (err) {
      if (err instanceof BotAuthError) {
        return new Response(err.message, { status: 401 })
      }
      throw err
    }

    const body = await request.text()
    const activity = JSON.parse(body) as CoreActivity

    if (!this.chat) {
      throw new Error(
        'TeamsAdapter has not been initialized. ' +
        'Call await chat.initialize() before handling webhooks, ' +
        'or pass the adapter to the Chat constructor so it initializes automatically.'
      )
    }

    const botId = this.app.options.clientId ?? process.env['CLIENT_ID'] ?? ''
    const threadId = encodeThreadId({
      serviceUrl: activity.serviceUrl,
      conversationId: activity.conversation.id,
    })

    if (activity.type === 'message') {
      this.chat.processMessage(
        this,
        threadId,
        () => Promise.resolve(activityToMessage(activity, threadId, botId)),
        options
      )
    } else if (activity.type === 'messageReaction') {
      const addedReactions = (activity.reactionsAdded ?? []).map((r) => ({ type: r.type, added: true }))
      const removedReactions = (activity.reactionsRemoved ?? []).map((r) => ({ type: r.type, added: false }))
      for (const reaction of [...addedReactions, ...removedReactions]) {
        this.chat.processReaction(
          {
            adapter: this,
            threadId,
            messageId: activity.replyToId ?? '',
            added: reaction.added,
            emoji: toEmojiValue(reaction.type),
            rawEmoji: reaction.type,
            user: accountToAuthor(activity.from, botId),
            raw: activity,
          },
          options
        )
      }
    } else if (activity.type === 'invoke' && activity.name === 'adaptiveCard/action') {
      const invokeValue = activity.value as Record<string, unknown> | undefined
      const actionPayload = invokeValue?.['action'] as Record<string, unknown> | undefined
      const actionId = actionPayload?.['id'] as string ?? ''
      const actionData = actionPayload?.['data'] as Record<string, unknown> | undefined
      this.chat.processAction(
        {
          adapter: this,
          threadId,
          messageId: activity.replyToId ?? activity.id ?? '',
          actionId,
          value: actionData != null ? JSON.stringify(actionData) : undefined,
          user: accountToAuthor(activity.from, botId),
          raw: activity,
        },
        options
      )
      // Bot Framework requires an invokeResponse for invoke activities
      return new Response(
        JSON.stringify({ type: 'invokeResponse', value: { status: 200 } }),
        { status: 200, headers: { 'Content-Type': 'application/json' } }
      )
    }

    return new Response('{}', {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    })
  }

  parseMessage (raw: CoreActivity): Message<CoreActivity> {
    const threadId = encodeThreadId({
      serviceUrl: raw.serviceUrl,
      conversationId: raw.conversation.id,
    })
    const botId = this.app.options.clientId ?? process.env['CLIENT_ID'] ?? ''
    return activityToMessage(raw, threadId, botId)
  }

  async postMessage (threadId: string, message: AdapterPostableMessage): Promise<RawMessage<CoreActivity>> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const baseConversationId = conversationId.split(';')[0]
    const activity = this.toActivity(message)
    const response = await this.app.conversationClient.sendActivityAsync(serviceUrl, baseConversationId, activity)
    const id = response?.id ?? ''
    return { id, threadId, raw: { ...activity, id, serviceUrl, channelId: 'msteams', from: { id: '' }, recipient: { id: '' }, conversation: { id: baseConversationId } } as CoreActivity }
  }

  async editMessage (threadId: string, messageId: string, message: AdapterPostableMessage): Promise<RawMessage<CoreActivity>> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const baseConversationId = conversationId.split(';')[0]
    const activity = this.toActivity(message)
    await this.app.conversationClient.updateActivityAsync(serviceUrl, baseConversationId, messageId, activity)
    return { id: messageId, threadId, raw: { ...activity, id: messageId, serviceUrl, channelId: 'msteams', from: { id: '' }, recipient: { id: '' }, conversation: { id: baseConversationId } } as CoreActivity }
  }

  async deleteMessage (threadId: string, messageId: string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const baseConversationId = conversationId.split(';')[0]
    await this.app.conversationClient.deleteActivityAsync(serviceUrl, baseConversationId, messageId)
  }

  async addReaction (threadId: string, messageId: string, emoji: EmojiValue | string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const baseConversationId = conversationId.split(';')[0]
    const reactionType = typeof emoji === 'string' ? emoji : emoji.name
    await this.app.conversationClient.addReactionAsync(serviceUrl, baseConversationId, messageId, reactionType)
  }

  async removeReaction (threadId: string, messageId: string, emoji: EmojiValue | string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const baseConversationId = conversationId.split(';')[0]
    const reactionType = typeof emoji === 'string' ? emoji : emoji.name
    await this.app.conversationClient.deleteReactionAsync(serviceUrl, baseConversationId, messageId, reactionType)
  }

  async fetchMessages (_threadId: string, _options?: FetchOptions): Promise<FetchResult<CoreActivity>> {
    throw new NotImplementedError('fetchMessages is not supported by the Bot Framework REST API')
  }

  async fetchThread (threadId: string): Promise<ThreadInfo> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const baseConversationId = conversationId.split(';')[0]
    const conv = await this.app.conversationClient.getConversationAccountAsync(serviceUrl, baseConversationId)
    return {
      id: threadId,
      channelId: this.channelIdFromThreadId(threadId),
      channelName: conv?.name,
      isDM: conv?.conversationType === 'personal',
      metadata: conv ? (conv as unknown as Record<string, unknown>) : {},
    }
  }

  // Teams has a single typing indicator state — _status is accepted but ignored.
  async startTyping (threadId: string, _status?: string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const baseConversationId = conversationId.split(';')[0]
    await this.app.conversationClient.sendActivityAsync(serviceUrl, baseConversationId, { type: 'typing' })
  }

  renderFormatted (content: FormattedContent): string {
    return this.converter.fromAst(content)
  }

  // ── Helpers ───────────────────────────────────────────────────────────────

  private toActivity (message: AdapterPostableMessage): Partial<CoreActivity> {
    // JSX CardElement or { card: CardElement } → Adaptive Card attachment
    if (isCardElement(message)) {
      return {
        type: 'message',
        attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: cardToAdaptiveCard(message) }],
      }
    }
    if (typeof message === 'object' && 'card' in message) {
      return {
        type: 'message',
        attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: cardToAdaptiveCard(message.card) }],
      }
    }

    // AST → markdown (Teams supports markdown natively)
    if (typeof message === 'object' && 'ast' in message) {
      const text = convertEmojiPlaceholders(stringifyMarkdown(message.ast), 'teams')
      return { type: 'message', text, textFormat: 'markdown' }
    }

    // Markdown → markdown
    if (typeof message === 'object' && 'markdown' in message) {
      const text = convertEmojiPlaceholders(message.markdown, 'teams')
      return { type: 'message', text, textFormat: 'markdown' }
    }

    // Raw HTML string
    if (typeof message === 'object' && 'raw' in message) {
      return { type: 'message', text: message.raw, textFormat: 'html' }
    }

    // Plain string
    const text = convertEmojiPlaceholders(message as string, 'teams')
    return { type: 'message', text, textFormat: 'plain' }
  }
}

// ── Message parsing ─────────────────────────────────────────────────────────

function activityToMessage (activity: CoreActivity, threadId: string, botId: string): Message<CoreActivity> {
  const rawText = activity.text ?? ''
  // Normalise Teams HTML to markdown for consistent mdast parsing
  const markdownText = activity.textFormat === 'html' ? teamsHtmlToMarkdown(rawText) : rawText

  // Detect @-mention via activity entities — more reliable than text matching
  // because the mention display name in Teams may differ from adapter.userName.
  // Teams prefixes bot IDs with "28:" in mention entities (e.g. "28:abc-123"),
  // while clientId is the bare GUID, so we check both exact and suffix match.
  const isMention = (activity.entities ?? []).some((e) => {
    if (e.type !== 'mention') return false
    const mentioned = (e as Record<string, unknown>)['mentioned'] as Record<string, unknown> | undefined
    const mentionedId = mentioned?.['id'] as string | undefined
    if (!mentionedId || !botId) return false
    return mentionedId === botId || mentionedId.endsWith(`:${botId}`)
  })

  const data: MessageData<CoreActivity> = {
    id: activity.id ?? '',
    threadId,
    text: markdownText,
    formatted: parseMarkdown(markdownText),
    raw: activity,
    author: accountToAuthor(activity.from, activity.recipient?.id ?? botId),
    metadata: {
      dateSent: activity.timestamp ? new Date(activity.timestamp) : new Date(),
      edited: false,
    },
    attachments: mapAttachments(activity.attachments ?? []),
    isMention,
  }

  return new Message(data)
}

function accountToAuthor (account: ChannelAccount, botId: string) {
  return {
    userId: account.id,
    userName: account.name ?? account.id,
    fullName: account.name ?? account.id,
    isBot: (account.role === 'bot' ? true : 'unknown') as boolean | 'unknown',
    // Teams sends bot IDs as "28:{appId}" in activity from/recipient fields
    isMe: account.id === botId || account.id.endsWith(`:${botId}`),
  }
}

type ChatAttachment = {
  type: 'image' | 'file' | 'video' | 'audio'
  url?: string
  name?: string
  mimeType?: string
  fetchData?: () => Promise<Buffer>
}

function mapAttachments (botAttachments: BotAttachment[]): ChatAttachment[] {
  const result: ChatAttachment[] = []
  for (const a of botAttachments) {
    // Skip Teams internal HTML-body attachments that lack a content URL
    if (a.contentType === 'text/html' && !a.contentUrl) continue
    const type = mimeToAttachmentType(a.contentType)
    if (type && a.contentUrl) {
      const url = a.contentUrl
      result.push({
        type,
        url,
        name: a.name,
        mimeType: a.contentType,
        fetchData: async () => {
          const res = await fetch(url)
          const buf = await res.arrayBuffer()
          return Buffer.from(buf)
        },
      })
    }
  }
  return result
}

function mimeToAttachmentType (contentType: string): 'image' | 'file' | 'video' | 'audio' | null {
  if (contentType.startsWith('image/')) return 'image'
  if (contentType.startsWith('video/')) return 'video'
  if (contentType.startsWith('audio/')) return 'audio'
  // Skip Adaptive Cards and Hero cards — they are not file attachments
  if (contentType.startsWith('application/vnd.microsoft.card')) return null
  return 'file'
}

/**
 * Convert a Teams reaction type string to a Chat SDK EmojiValue.
 * Falls back to an ad-hoc EmojiValue if the name is not in the standard map.
 */
function toEmojiValue (reactionType: string): EmojiValue {
  // Map Teams reaction names to Chat SDK normalized emoji names
  const teamsToEmoji: Record<string, string> = {
    like: 'thumbs_up',
    heart: 'heart',
    laugh: 'laughing',
    surprised: 'open_mouth',
    sad: 'cry',
    angry: 'rage',
  }
  const name = teamsToEmoji[reactionType] ?? reactionType
  try {
    return getEmoji(name)
  } catch {
    // Return a minimal EmojiValue for unknown reaction types
    return {
      name,
      toJSON: () => `{{emoji:${name}}}`,
      toString: () => `{{emoji:${name}}}`,
    }
  }
}
