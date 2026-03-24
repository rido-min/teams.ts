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

  private chat!: ChatInstance
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
   * Handle an incoming Bot Framework webhook request.
   *
   * Validates the Bearer token, dispatches message activities to the Chat
   * SDK, and returns a `200 {}` response immediately (fire-and-forget via
   * `waitUntil`).
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

    if (activity.type === 'message') {
      const threadId = encodeThreadId({
        serviceUrl: activity.serviceUrl,
        conversationId: activity.conversation.id,
      })
      this.chat.processMessage(
        this,
        threadId,
        () => Promise.resolve(this.parseMessage(activity)),
        options
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
    const activity = this.toActivity(message)
    const response = await this.app.conversationClient.sendActivityAsync(serviceUrl, conversationId, activity)
    const id = response?.id ?? ''
    return { id, threadId, raw: { ...activity, id, serviceUrl, channelId: 'msteams', from: { id: '' }, recipient: { id: '' }, conversation: { id: conversationId } } as CoreActivity }
  }

  async editMessage (threadId: string, messageId: string, message: AdapterPostableMessage): Promise<RawMessage<CoreActivity>> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const activity = this.toActivity(message)
    await this.app.conversationClient.updateActivityAsync(serviceUrl, conversationId, messageId, activity)
    return { id: messageId, threadId, raw: { ...activity, id: messageId, serviceUrl, channelId: 'msteams', from: { id: '' }, recipient: { id: '' }, conversation: { id: conversationId } } as CoreActivity }
  }

  async deleteMessage (threadId: string, messageId: string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    await this.app.conversationClient.deleteActivityAsync(serviceUrl, conversationId, messageId)
  }

  async addReaction (threadId: string, messageId: string, emoji: EmojiValue | string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const reactionType = typeof emoji === 'string' ? emoji : emoji.name
    await this.app.conversationClient.addReactionAsync(serviceUrl, conversationId, messageId, reactionType)
  }

  async removeReaction (threadId: string, messageId: string, emoji: EmojiValue | string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const reactionType = typeof emoji === 'string' ? emoji : emoji.name
    await this.app.conversationClient.deleteReactionAsync(serviceUrl, conversationId, messageId, reactionType)
  }

  async fetchMessages (_threadId: string, _options?: FetchOptions): Promise<FetchResult<CoreActivity>> {
    throw new NotImplementedError('fetchMessages is not supported by the Bot Framework REST API')
  }

  async fetchThread (threadId: string): Promise<ThreadInfo> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    const conv = await this.app.conversationClient.getConversationAccountAsync(serviceUrl, conversationId)
    return {
      id: threadId,
      channelId: this.channelIdFromThreadId(threadId),
      channelName: conv?.name,
      isDM: conv?.conversationType === 'personal',
      metadata: conv ? (conv as unknown as Record<string, unknown>) : {},
    }
  }

  async startTyping (threadId: string): Promise<void> {
    const { serviceUrl, conversationId } = decodeThreadId(threadId)
    await this.app.conversationClient.sendActivityAsync(serviceUrl, conversationId, { type: 'typing' })
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
        attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: message }],
      }
    }
    if (typeof message === 'object' && 'card' in message) {
      return {
        type: 'message',
        attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: message.card }],
      }
    }

    // AST → markdown (Teams supports markdown natively)
    if (typeof message === 'object' && 'ast' in message) {
      return { type: 'message', text: stringifyMarkdown(message.ast), textFormat: 'markdown' }
    }

    // Markdown → markdown
    if (typeof message === 'object' && 'markdown' in message) {
      return { type: 'message', text: message.markdown, textFormat: 'markdown' }
    }

    // Raw HTML string
    if (typeof message === 'object' && 'raw' in message) {
      return { type: 'message', text: message.raw, textFormat: 'html' }
    }

    // Plain string
    return { type: 'message', text: message as string, textFormat: 'plain' }
  }
}

// ── Message parsing ─────────────────────────────────────────────────────────

function activityToMessage (activity: CoreActivity, threadId: string, botId: string): Message<CoreActivity> {
  const rawText = activity.text ?? ''
  // Normalise Teams HTML to markdown for consistent mdast parsing
  const markdownText = activity.textFormat === 'html' ? teamsHtmlToMarkdown(rawText) : rawText

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
  }

  return new Message(data)
}

function accountToAuthor (account: ChannelAccount, botId: string) {
  return {
    userId: account.id,
    userName: account.name ?? account.id,
    fullName: account.name ?? account.id,
    isBot: (account.role === 'bot' ? true : 'unknown') as boolean | 'unknown',
    isMe: account.id === botId,
  }
}

type ChatAttachment = {
  type: 'image' | 'file' | 'video' | 'audio'
  url?: string
  name?: string
  mimeType?: string
}

function mapAttachments (botAttachments: BotAttachment[]): ChatAttachment[] {
  const result: ChatAttachment[] = []
  for (const a of botAttachments) {
    const type = mimeToAttachmentType(a.contentType)
    if (type && a.contentUrl) {
      result.push({ type, url: a.contentUrl, name: a.name, mimeType: a.contentType })
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
