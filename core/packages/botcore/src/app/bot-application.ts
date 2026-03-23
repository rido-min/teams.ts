// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { IncomingMessage, ServerResponse } from 'node:http'
import type { CoreActivity, ResourceResponse } from '../schema/core-activity.js'
import type { ITurnMiddleware } from '../middleware/i-turn-middleware.js'
import { ConversationClient } from '../clients/conversation-client.js'
import { UserTokenClient } from '../clients/user-token-client.js'
import { TokenManager } from '../auth/token-manager.js'
import type { BotApplicationOptions } from './bot-application-options.js'
import { type ActivityHandler, createContext } from './activity-context.js'

/**
 * Core bot application that processes incoming Bot Framework activities.
 *
 * `BotApplication` is web-server-agnostic — use {@link processAsync} with
 * Express (or any Node.js `http.Server`) and {@link processBody} with
 * frameworks that manage their own HTTP response lifecycle (e.g. Hono).
 *
 * @example
 * ```ts
 * const app = new BotApplication({ clientId: '...', clientSecret: '...' });
 *
 * app.on('message', async ({ activity, send }) => {
 *   await send(`you said "${activity.text}"`);
 * });
 * ```
 */
export class BotApplication {
  /** Resolved options for this application instance. */
  readonly options: BotApplicationOptions

  /** Client for sending, updating, and deleting activities via the Bot Framework API. */
  readonly conversationClient: ConversationClient

  /** Client for OAuth token operations via the Bot Framework token service. */
  readonly userTokenClient: UserTokenClient

  private readonly middlewares: ITurnMiddleware[] = []
  private readonly handlers = new Map<string, ActivityHandler>()
  private readonly tokenManager: TokenManager

  /**
   * @param options - Application configuration. All fields are optional and
   *   fall back to environment variables (`CLIENT_ID`, `CLIENT_SECRET`, etc.).
   */
  constructor (options: BotApplicationOptions = {}) {
    this.options = options
    this.tokenManager = new TokenManager(options)
    const tokenProvider = () =>
      this.tokenManager.getBotToken().then((t) => {
        if (!t) throw new Error('No credentials configured — set CLIENT_ID and CLIENT_SECRET (or use managed identity)')
        return t
      })
    this.conversationClient = new ConversationClient(tokenProvider)
    this.userTokenClient = new UserTokenClient(tokenProvider)
  }

  /**
   * Register a handler for a given activity type.
   *
   * Only one handler per type is supported; registering the same type twice
   * replaces the previous handler.
   *
   * @param type - Activity type string (e.g. `"message"`, `"conversationUpdate"`).
   * @param handler - Async function called with the turn context.
   * @returns `this` for method chaining.
   *
   * @example
   * app.on('message', async ({ activity, send }) => {
   *   await send(`you said "${activity.text}"`);
   * });
   */
  on (type: string, handler: ActivityHandler): this {
    this.handlers.set(type, handler)
    return this
  }

  /**
   * Add a middleware to the turn processing pipeline.
   *
   * Middleware runs in registration order before the activity handler.
   *
   * @param middleware - Middleware implementing {@link ITurnMiddleware}.
   * @returns `this` for method chaining.
   */
  use (middleware: ITurnMiddleware): this {
    this.middlewares.push(middleware)
    return this
  }

  /**
   * Process an incoming HTTP request (Express / Node.js `http.Server`).
   *
   * Reads the request body, validates the activity, runs the middleware
   * pipeline, and writes a `200 {}` response. Writes `500` on error.
   *
   * @param req - Node.js `IncomingMessage`.
   * @param res - Node.js `ServerResponse`.
   */
  async processAsync (req: IncomingMessage, res: ServerResponse): Promise<void> {
    try {
      const body = await readBody(req)
      await this.processBody(body)
      res.writeHead(200, { 'Content-Type': 'application/json' })
      res.end('{}')
    } catch {
      res.writeHead(500)
      res.end('Internal server error')
    }
  }

  /**
   * Process a raw JSON request body without touching the HTTP layer.
   *
   * Use this with frameworks that manage their own response lifecycle (e.g.
   * Hono), where writing to `ServerResponse` directly would cause
   * `ERR_HTTP_HEADERS_SENT`.
   *
   * @param body - Raw JSON string of the activity payload.
   */
  async processBody (body: string): Promise<void> {
    const activity = JSON.parse(body) as CoreActivity
    assertActivity(activity)
    await this.runPipelineAsync(activity)
  }

  /**
   * Proactively send an activity to a conversation.
   *
   * @param serviceUrl - Bot Framework service URL.
   * @param conversationId - Target conversation ID.
   * @param activity - Activity payload to send.
   * @returns Resource response with the new activity ID.
   */
  async sendActivityAsync (
    serviceUrl: string,
    conversationId: string,
    activity: Partial<CoreActivity>
  ): Promise<ResourceResponse | undefined> {
    return this.conversationClient.sendActivityAsync(serviceUrl, conversationId, activity)
  }

  /** @internal Dispatch the activity to its registered handler. */
  protected async handleActivityAsync (activity: CoreActivity): Promise<void> {
    const handler = this.handlers.get(activity.type)
    if (handler) {
      await handler(createContext(activity, this))
    }
  }

  private async runPipelineAsync (activity: CoreActivity): Promise<void> {
    let index = 0
    const next = async (): Promise<void> => {
      if (index < this.middlewares.length) {
        await this.middlewares[index++].onTurnAsync(this, activity, next)
      } else {
        await this.handleActivityAsync(activity)
      }
    }
    await next()
  }
}

function assertActivity (value: unknown): asserts value is CoreActivity {
  if (typeof value !== 'object' || value === null) {
    throw new Error('Activity must be a JSON object')
  }
  const a = value as Record<string, unknown>
  if (typeof a['type'] !== 'string' || !a['type']) {
    throw new Error('Activity missing required field: type')
  }
  if (typeof a['serviceUrl'] !== 'string' || !a['serviceUrl']) {
    throw new Error('Activity missing required field: serviceUrl')
  }
  if (
    typeof a['conversation'] !== 'object' ||
    a['conversation'] === null ||
    typeof (a['conversation'] as Record<string, unknown>)['id'] !== 'string' ||
    !(a['conversation'] as Record<string, unknown>)['id']
  ) {
    throw new Error('Activity missing required field: conversation.id')
  }
}

function readBody (req: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    req.on('data', (chunk: Buffer) => chunks.push(chunk))
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    req.on('error', reject)
  })
}
