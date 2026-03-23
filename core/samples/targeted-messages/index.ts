// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// Sample: Targeted Messages
//
// Targeted Messages allow a bot to send a message visible only to a specific
// user in a group conversation, without revealing it to other participants.
//
// This sample demonstrates the full lifecycle within a single message handler:
//   1. Send a targeted message (only visible to the sender)
//   2. Update it in-place
//   3. Delete it
//
// Run: npx tsx index.ts

import express from 'express'
import { BotApplication, botAuthExpress } from '@microsoft/teams.botcore'

// ── Bot ───────────────────────────────────────────────────────────────────────

const app = new BotApplication()

app.on('message', async ({ activity }) => {
  const { serviceUrl, conversation, from } = activity

  // The recipient must have isTargeted: true so the Bot Framework appends
  // ?isTargetedActivity=true — Teams then delivers the message only to `from`.
  const targetedRecipient = { ...from, isTargeted: true }

  // 1. Send a targeted message visible only to the user who sent the message.
  const response = await app.conversationClient.sendActivityAsync(
    serviceUrl,
    conversation.id,
    {
      type: 'message',
      text: `Only you can see this, ${from.name ?? from.id}.`,
      recipient: targetedRecipient,
    }
  )

  if (!response) return
  const activityId = response.id

  // 2. Update the targeted message in-place.
  // Do not include `recipient` in the update body — Teams rejects it with
  // "Cannot edit Recipient of Targeted Message".
  await app.conversationClient.updateTargetedActivityAsync(
    serviceUrl,
    conversation.id,
    activityId,
    {
      type: 'message',
      text: `Updated: only you can still see this, ${from.name ?? from.id}.`,
    }
  )

  // 3. Delete the targeted message.
  await app.conversationClient.deleteTargetedActivityAsync(
    serviceUrl,
    conversation.id,
    activityId
  )
})

// ── Server ────────────────────────────────────────────────────────────────────

const server = express()

server.post('/api/messages', botAuthExpress(), (req, res) => {
  app.processAsync(req, res)
})

server.get('/health', (_req, res) => res.json({ status: 'ok' }))

const PORT = Number(process.env.PORT ?? 3978)
server.listen(PORT, () => {
  console.log(`Listening on http://localhost:${PORT}/api/messages`)
})
