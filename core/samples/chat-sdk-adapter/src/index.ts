// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// Sample: Teams bot using the Chat SDK adapter
//
// Demonstrates how to use @microsoft/teams.botcore-chat-adapter to connect
// a Chat SDK bot to Microsoft Teams via Hono.
//
// Run: npx tsx index.ts

import { serve } from '@hono/node-server'
import { Chat } from 'chat'
import { createMemoryState } from '@chat-adapter/state-memory'
import { Hono } from 'hono'
import { TeamsAdapter } from '@microsoft/teams.botcore-chat-adapter'

// ── Shared command handler ────────────────────────────────────────────────────

import type { Thread, Message } from 'chat'

// ── Adapter & Chat instance ───────────────────────────────────────────────────

// Credentials are auto-detected from CLIENT_ID / CLIENT_SECRET / TENANT_ID env vars.
const adapter = new TeamsAdapter({ userName: 'TeamsBot' })

const chat = new Chat({
  adapters: { teams: adapter },
  // In production replace createMemoryState() with a Redis- or DB-backed adapter.
  state: createMemoryState(),
  userName: 'TeamsBot',
})

// ── Handlers ──────────────────────────────────────────────────────────────────

// Called when the bot is @-mentioned in a conversation it isn't yet subscribed to.
// Subscribe to the thread so follow-up messages also reach onSubscribedMessage.
chat.onNewMention(async (thread, message) => {
  console.log(`[mention] from=${message.author.fullName} text="${message.text}"`)
  await thread.subscribe()
  await handleMessage(thread, message)
})

// Called for every message in a personal (1:1) chat.
chat.onDirectMessage(async (thread, message) => {
  console.log(`[dm] from=${message.author.fullName} text="${message.text}"`)
  await thread.subscribe()
  await handleMessage(thread, message)
})

// Called for all follow-up messages in subscribed threads.
chat.onSubscribedMessage(async (thread, message) => {
  console.log(`[subscribed] from=${message.author.fullName} text="${message.text}"`)
  await handleMessage(thread, message)
})

async function handleMessage (thread: Thread, message: Message): Promise<void> {
  if (message.text.trim() === '') return

  // Show a typing indicator while "processing".
  await thread.startTyping()

  const lower = message.text.toLowerCase().trim()

  if (lower === 'help') {
    // Reply with a markdown-formatted help card.
    await thread.post({
      markdown: [
        '**Available commands**',
        '',
        '| Command | Description |',
        '|---------|-------------|',
        '| `help` | Show this message |',
        '| `ping` | Get a pong reply |',
        '| `echo <text>` | Echo text back |',
        '| `react` | React to your message |',
      ].join('\n'),
    })
    return
  }

  if (lower === 'ping') {
    await thread.post('Pong!')
    return
  }

  if (lower.startsWith('echo ')) {
    const payload = message.text.slice(5)
    await thread.post({ markdown: `**Echo:** ${payload}` })
    return
  }

  if (lower === 'react') {
    // Use the adapter's public addReaction to react to the user's message.
    await adapter.addReaction(message.threadId, message.id, 'like')
    await thread.post('Reacted to your message!')
    return
  }

  // Default: plain echo.
  await thread.post(`You said: "${message.text}"`)
}

// ── HTTP server (Hono + Node.js adapter) ─────────────────────────────────────

const honoApp = new Hono()

// TeamsAdapter.handleWebhook() validates the Bot Framework JWT, dispatches the
// activity to the Chat SDK, and returns a 200 immediately.
honoApp.post('/api/messages', (c) => {
  return adapter.handleWebhook(c.req.raw, {
    waitUntil: (p) => p.catch(console.error),
  })
})

honoApp.get('/health', (c) => c.json({ status: 'ok' }))

const PORT = Number(process.env['PORT'] ?? 3978)
serve({ fetch: honoApp.fetch, port: PORT }, () => {
  console.log(`Bot listening on http://localhost:${PORT}/api/messages`)
})
