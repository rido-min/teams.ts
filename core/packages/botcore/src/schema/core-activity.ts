// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/** Represents a user or bot account in a channel. */
export interface ChannelAccount {
  /** Unique identifier for the account. */
  id: string;
  /** Display name of the account. */
  name?: string;
  /** Azure Active Directory object ID. */
  aadObjectId?: string;
  /** Role of the entity (e.g. `"user"`, `"bot"`, `"skill"`). */
  role?: string;
  /** Additional channel-specific properties. */
  properties?: Record<string, unknown>;
}

/** Represents a conversation in a channel. */
export interface ConversationAccount {
  /** Unique identifier for the conversation. */
  id: string;
  /** Display name of the conversation. */
  name?: string;
  /** Whether the conversation contains more than two participants. */
  isGroup?: boolean;
  /** Type of conversation (e.g. `"personal"`, `"channel"`). */
  conversationType?: string;
  /** Tenant ID the conversation belongs to. */
  tenantId?: string;
  /** Additional channel-specific properties. */
  properties?: Record<string, unknown>;
}

/** An entity associated with an activity (e.g. a mention). */
export interface Entity {
  /** Entity type (e.g. `"mention"`, `"clientInfo"`). */
  type: string;
  [key: string]: unknown;
}

/** A file or card attachment on an activity. */
export interface Attachment {
  /** MIME type of the attachment content. */
  contentType: string;
  /** URL to retrieve the attachment content. */
  contentUrl?: string;
  /** Embedded content (e.g. an Adaptive Card JSON object). */
  content?: unknown;
  /** Filename or display name. */
  name?: string;
  /** URL to a thumbnail image. */
  thumbnailUrl?: string;
}

/** A reaction (emoji) applied to a message. */
export interface MessageReaction {
  /** Reaction type string (e.g. `"like"`, `"heart"`). */
  type: string;
}

/**
 * The base activity payload received from the Bot Framework.
 * All channel-specific activity types extend or narrow this interface.
 */
export interface CoreActivity {
  /** Activity type (e.g. `"message"`, `"conversationUpdate"`). */
  type: string;
  /** Activity ID assigned by the channel. */
  id?: string;
  /** Channel identifier (e.g. `"msteams"`, `"directline"`). */
  channelId: string;
  /** Service URL for the sending channel — used to reply. */
  serviceUrl: string;
  /** Account that sent the activity. */
  from: ChannelAccount;
  /** Account that received the activity (the bot). */
  recipient: ChannelAccount;
  /** Conversation this activity belongs to. */
  conversation: ConversationAccount;
  /** Channel-specific payload (e.g. Teams channel data). */
  channelData?: unknown;
  /** Structured entities attached to the activity. */
  entities?: Entity[];
  /** File or card attachments. */
  attachments?: Attachment[];
  /** Activity value payload (used by invoke and event activities). */
  value?: unknown;
  /** ID of the activity this is a reply to. */
  replyToId?: string;
  /** Text content of a message activity. */
  text?: string;
  /** Name of the event or invoke operation. */
  name?: string;
  /** Locale of the activity (e.g. `"en-US"`). */
  locale?: string;
  /** UTC timestamp when the activity was sent. */
  timestamp?: string;
  /** Local timestamp when the activity was sent. */
  localTimestamp?: string;
  /** Reactions added to a message (messageReaction activities). */
  reactionsAdded?: MessageReaction[];
  /** Reactions removed from a message (messageReaction activities). */
  reactionsRemoved?: MessageReaction[];
  /** Members added to the conversation (conversationUpdate activities). */
  membersAdded?: ChannelAccount[];
  /** Members removed from the conversation (conversationUpdate activities). */
  membersRemoved?: ChannelAccount[];
  /** Additional custom properties. */
  properties?: Record<string, unknown>;
}

/** Response returned after creating or sending an activity. */
export interface ResourceResponse {
  /** ID of the created resource. */
  id: string;
}

/** Response returned after creating a new conversation. */
export interface ConversationResourceResponse extends ResourceResponse {
  /** Service URL for the new conversation. */
  serviceUrl: string;
  /** ID of the activity that bootstrapped the conversation. */
  activityId: string;
}

/**
 * A page of members with an optional continuation token for subsequent pages.
 * @typeParam T - Member type (e.g. `ChannelAccount`).
 */
export interface PagedMembersResult<T> {
  /** Members on this page. */
  members: T[];
  /** Opaque token to pass on the next call to retrieve the following page. */
  continuationToken?: string;
}

/** Parameters used to create a new conversation. */
export interface ConversationParameters {
  /** Whether to create a group conversation. */
  isGroup?: boolean;
  /** Bot account to include in the conversation. */
  bot?: ChannelAccount;
  /** Initial member list. */
  members?: ChannelAccount[];
  /** Optional topic name. */
  topicName?: string;
  /** Tenant ID. */
  tenantId?: string;
  /** Initial activity to post when the conversation is created. */
  activity?: Partial<CoreActivity>;
  /** Channel-specific data. */
  channelData?: unknown;
}

/** A page of conversations with an optional continuation token. */
export interface ConversationsResult {
  /** Conversations on this page. */
  conversations: ConversationAccount[];
  /** Opaque token to pass on the next call to retrieve the following page. */
  continuationToken?: string;
}

/** A transcript (ordered list of activities) for a conversation. */
export interface Transcript {
  /** Activities in the transcript, ordered by timestamp. */
  activities: CoreActivity[];
}
