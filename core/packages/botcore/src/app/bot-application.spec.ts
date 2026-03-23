// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, it } from 'node:test'
import assert from 'node:assert/strict'
import type { IncomingMessage, ServerResponse } from 'node:http'
import { BotApplication } from './bot-application.js'
import type { CoreActivity } from '../schema/core-activity.js'
import type { ITurnMiddleware, NextTurn } from '../middleware/i-turn-middleware.js'

function makeRequest (body: unknown): IncomingMessage {
  const json = JSON.stringify(body)
  let consumed = false
  const req = {
    headers: {},
    on (event: string, handler: (chunk?: Buffer) => void) {
      if (event === 'data' && !consumed) {
        consumed = true
        handler(Buffer.from(json, 'utf8'))
      } else if (event === 'end') {
        handler()
      }
      return req
    },
  } as unknown as IncomingMessage
  return req
}

function makeResponse (): { res: ServerResponse; state: { statusCode: number; body: string } } {
  const state = { statusCode: 0, body: '' }
  const res = {
    writeHead (code: number) { state.statusCode = code },
    end (data?: string) { state.body = data ?? '' },
  } as unknown as ServerResponse
  return { res, state }
}

const sampleActivity: CoreActivity = {
  type: 'message',
  id: 'act-1',
  channelId: 'msteams',
  serviceUrl: 'https://smba.trafficmanager.net/',
  from: { id: 'user-1', name: 'User' },
  recipient: { id: 'bot-1', name: 'Bot' },
  conversation: { id: 'conv-1' },
  text: 'Hello',
}

describe('BotApplication', () => {
  it('calls a registered handler for the matching activity type', async () => {
    const app = new BotApplication()
    let received: CoreActivity | undefined

    app.on('message', async ({ activity }) => {
      received = activity
    })

    const { res, state } = makeResponse()
    await app.processAsync(makeRequest(sampleActivity), res)

    assert.equal(state.statusCode, 200)
    assert.equal(received?.type, 'message')
    assert.equal(received?.text, 'Hello')
  })

  it('does not call handler for non-matching activity type', async () => {
    const app = new BotApplication()
    let called = false

    app.on('event', async () => { called = true })

    const { res } = makeResponse()
    await app.processAsync(makeRequest(sampleActivity), res) // type is 'message'

    assert.equal(called, false)
  })

  it('send helper replies to the current conversation', async () => {
    const app = new BotApplication()
    const sent: Array<{ serviceUrl: string; conversationId: string }> = []

    // Intercept outgoing calls
    app.conversationClient.sendActivityAsync = async (serviceUrl, conversationId) => {
      sent.push({ serviceUrl, conversationId })
      return { id: 'reply-1' }
    }

    app.on('message', async ({ send }) => {
      await send('pong')
    })

    const { res } = makeResponse()
    await app.processAsync(makeRequest(sampleActivity), res)

    assert.equal(sent.length, 1)
    assert.equal(sent[0].serviceUrl, sampleActivity.serviceUrl)
    assert.equal(sent[0].conversationId, sampleActivity.conversation.id)
  })

  it('on() returns this for chaining', () => {
    const app = new BotApplication()
    const result = app.on('message', async () => {})
    assert.equal(result, app)
  })

  it('runs middleware chain before the handler', async () => {
    const app = new BotApplication()
    const order: string[] = []

    const mw: ITurnMiddleware = {
      async onTurnAsync (_a, _act, next: NextTurn) {
        order.push('middleware')
        await next()
      },
    }
    app.use(mw)
    app.on('message', async () => { order.push('handler') })

    const { res } = makeResponse()
    await app.processAsync(makeRequest(sampleActivity), res)

    assert.deepEqual(order, ['middleware', 'handler'])
  })

  it('returns 200 even when no handler is registered', async () => {
    const app = new BotApplication()
    const { res, state } = makeResponse()
    await app.processAsync(makeRequest(sampleActivity), res)
    assert.equal(state.statusCode, 200)
  })

  it('returns 500 when activity is missing type', async () => {
    const app = new BotApplication()
    const bad = { ...sampleActivity, type: undefined }
    const { res, state } = makeResponse()
    await app.processAsync(makeRequest(bad), res)
    assert.equal(state.statusCode, 500)
  })

  it('returns 500 when activity is missing serviceUrl', async () => {
    const app = new BotApplication()
    const bad = { ...sampleActivity, serviceUrl: undefined }
    const { res, state } = makeResponse()
    await app.processAsync(makeRequest(bad), res)
    assert.equal(state.statusCode, 500)
  })

  it('returns 500 when activity is missing conversation.id', async () => {
    const app = new BotApplication()
    const bad = { ...sampleActivity, conversation: { id: '' } }
    const { res, state } = makeResponse()
    await app.processAsync(makeRequest(bad), res)
    assert.equal(state.statusCode, 500)
  })
})
