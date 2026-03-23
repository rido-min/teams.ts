// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// Sample: botcore with Hono (Node.js adapter)
// Run: npx tsx index.ts

import { Hono } from 'hono';
import { serve } from '@hono/node-server';
import { BotApplication, botAuthHono } from '@microsoft/teams.botcore';

// ── Bot ───────────────────────────────────────────────────────────────────────

// Credentials are auto-detected from CLIENT_ID / CLIENT_SECRET / TENANT_ID env vars.
const app = new BotApplication();

app.on('message', async ({ activity, send }) => {
  await send(`you said "${activity.text}"`);
});

app.on('conversationUpdate', async ({ activity }) => {
  console.log('conversation update', activity.membersAdded);
});

// ── Server ────────────────────────────────────────────────────────────────────

const honoApp = new Hono();

honoApp.post('/api/messages', botAuthHono(), async (c) => {
  const body = await c.req.text();
  await app.processBody(body);
  return c.json({});
});

honoApp.get('/health', (c) => c.json({ status: 'ok' }));

const PORT = Number(process.env.PORT ?? 3978);
serve({ fetch: honoApp.fetch, port: PORT }, () => {
  console.log(`Listening on http://localhost:${PORT}/api/messages`);
});
