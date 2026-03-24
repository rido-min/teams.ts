# Sample: Teams bot with the Chat SDK adapter

This sample shows how to connect a [Chat SDK](https://chat-sdk.dev) bot to
Microsoft Teams using `@microsoft/teams.botcore-chat-adapter`.

The adapter bridges the Chat SDK's platform-agnostic API with the Bot
Framework REST API provided by `@microsoft/teams.botcore`.

## What it demonstrates

| Feature | How |
|---------|-----|
| Receive @-mentions | `chat.onNewMention()` |
| Receive DMs | `chat.onDirectMessage()` |
| Receive follow-up messages | `chat.onSubscribedMessage()` |
| Send plain text | `thread.post("hello")` |
| Send markdown | `thread.post({ markdown: "**bold**" })` |
| Send Adaptive Cards | `thread.post(Card({ title: "...", children: [...] }))` |
| Card buttons (Action.Submit) | `chat.onAction("approve", handler)` |
| Card links (Action.OpenUrl) | `LinkButton({ url, label })` in `Actions([...])` |
| Reaction events | `chat.onReaction(handler)` |
| Show typing indicator | `thread.startTyping()` |
| Add a reaction | `adapter.addReaction(threadId, messageId, "like")` |
| Emoji placeholders | `emoji.thumbs_up` → 👍 in Teams |
| Webhook auth | JWT validation inside `TeamsAdapter.handleWebhook()` |

## Prerequisites

- Node.js 20+
- An Azure Bot registration (App ID + client secret or managed identity)
- The bot endpoint registered as `https://<your-host>/api/messages`

## Running locally

1. Install dependencies from the repo root:
   ```sh
   npm install
   ```

2. Set credentials as environment variables:
   ```sh
   export CLIENT_ID=<your-app-id>
   export CLIENT_SECRET=<your-client-secret>
   # export TENANT_ID=<tenant>   # optional
   # export PORT=3978            # optional, defaults to 3978
   ```

3. Start the bot:
   ```sh
   cd samples/chat-sdk-adapter
   npm start
   ```

4. Expose port `3978` with a tunnelling tool (e.g.
   [dev tunnels](https://learn.microsoft.com/azure/developer/dev-tunnels/overview)
   or ngrok) and point your Azure Bot messaging endpoint to
   `https://<tunnel>/api/messages`.

## Commands

Once the bot is installed in a Teams chat, try these messages:

| Message | Response |
|---------|----------|
| `help` | Markdown table of available commands |
| `ping` | `Pong! 🏓` |
| `echo <text>` | Bold echo of `<text>` |
| `card` | Adaptive Card with Approve / Reject buttons |
| `react` | Thumbs-up reaction on the user's message |
| anything else | Plain echo |

## Architecture

```
Teams ──HTTP──► Hono ──► adapter.handleWebhook()
                              │
                    validateBotToken()  ← Bot Framework JWT
                              │
                         activity.type?
                      ┌──────┼────────────────────────┐
                  message  messageReaction          invoke
                      │        │                      │
              processMessage  processReaction   processAction
                      │                               │
              ┌───────┼────────────────┐         onAction()
      onNewMention  onDirectMessage  onSubscribedMessage
              └───────┼────────────────┘
                 handleMessage()
                      │
                 thread.post() ──► ConversationClient ──► Teams
```

`TeamsAdapter` wraps `BotApplication` from `@microsoft/teams.botcore`.
Thread IDs encode both the `serviceUrl` and `conversationId` required by the
Bot Framework REST API.

## Configuration

```ts
const adapter = new TeamsAdapter({
  clientId: 'your-app-id',       // or CLIENT_ID env var
  clientSecret: 'your-secret',   // or CLIENT_SECRET env var
  userName: 'MyBot',             // bot display name
})

const chat = new Chat({
  adapters: { teams: adapter },
  state: createMemoryState(),    // swap for Redis/DB in production
  userName: 'MyBot',
})
```

For managed identity deployments omit `clientSecret` and set
`managedIdentityClientId` instead.

## Adaptive Cards

Send rich cards with buttons using the Chat SDK's card primitives:

```ts
import { Card, Text, Actions, Button, LinkButton, Fields, Field } from 'chat'

await thread.post(
  Card({
    title: 'Approval Request',
    children: [
      Text('Please review and approve or reject.'),
      Fields([
        Field({ label: 'Requested by', value: message.author.fullName }),
      ]),
      Actions([
        Button({ id: 'approve', label: 'Approve', style: 'primary' }),
        Button({ id: 'reject',  label: 'Reject',  style: 'danger'  }),
        LinkButton({ label: 'Docs', url: 'https://example.com' }),
      ]),
    ],
  })
)
```

Handle button clicks with `chat.onAction()`:

```ts
chat.onAction('approve', async (event) => {
  await event.thread?.post(`Approved by ${event.user.fullName}!`)
})
```

## Reaction events

```ts
chat.onReaction(async (event) => {
  console.log(`${event.user.userName} ${event.added ? 'added' : 'removed'} ${event.emoji}`)
})
```

Teams reaction types (`like`, `heart`, `laugh`, `surprised`, `sad`, `angry`)
are automatically mapped to Chat SDK emoji values.
