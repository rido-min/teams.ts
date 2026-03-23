// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// Sample: botcore with Express
// Run: npx tsx index.ts

import express from 'express';
import { BotApplication, botAuthExpress } from '@microsoft/teams.botcore';

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

const server = express();
// server.use(express.json());

server.post('/api/messages', botAuthExpress(), (req, res) => {
  console.log('received message', req.body);
  app.processAsync(req, res);
});

server.get('/health', (_req, res) => res.json({ status: 'ok' }));

const PORT = Number(process.env.PORT ?? 3978);
server.listen(PORT, () => {
  console.log(`Listening on http://localhost:${PORT}/api/messages`);
});
