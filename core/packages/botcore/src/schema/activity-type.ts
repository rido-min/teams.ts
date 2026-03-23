// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Well-known Bot Framework activity type strings.
 *
 * @example
 * app.on(ActivityType.Message, async ({ activity, send }) => {
 *   await send(`you said "${activity.text}"`);
 * });
 */
export const ActivityType = {
  Message: 'message',
  Typing: 'typing',
  Event: 'event',
  Invoke: 'invoke',
  InvokeResponse: 'invokeResponse',
  ConversationUpdate: 'conversationUpdate',
  MessageUpdate: 'messageUpdate',
  MessageDelete: 'messageDelete',
  MessageReaction: 'messageReaction',
  InstallationUpdate: 'installationUpdate',
  HandOff: 'handoff',
  Trace: 'trace',
  EndOfConversation: 'endOfConversation',
  Command: 'command',
  CommandResult: 'commandResult',
} as const

/** Union type of all known activity type strings. */
// eslint-disable-next-line @typescript-eslint/no-redeclare
export type ActivityType = (typeof ActivityType)[keyof typeof ActivityType]
