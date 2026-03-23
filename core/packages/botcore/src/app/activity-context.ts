// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { CoreActivity, ResourceResponse } from '../schema/core-activity.js';
import type { BotApplication } from './bot-application.js';

/**
 * The turn context passed to every activity handler.
 *
 * @typeParam TActivity - Narrowed activity type for the specific handler.
 */
export interface ActivityContext<TActivity extends CoreActivity = CoreActivity> {
  /** The incoming activity. */
  readonly activity: TActivity;

  /**
   * Send a reply to the current conversation.
   * Pass a string for a plain text message, or a partial activity for full control.
   */
  send(text: string): Promise<ResourceResponse | undefined>;
  send(activity: Partial<CoreActivity>): Promise<ResourceResponse | undefined>;
}

/**
 * A function that handles a specific activity type.
 *
 * @typeParam TActivity - Narrowed activity type expected by this handler.
 */
export type ActivityHandler<TActivity extends CoreActivity = CoreActivity> = (
  ctx: ActivityContext<TActivity>
) => Promise<void>;

/**
 * Create an {@link ActivityContext} for the given activity and application.
 *
 * @param activity - The incoming activity.
 * @param app - The `BotApplication` instance processing the turn.
 * @returns A context object bound to the activity's conversation.
 */
export function createContext<TActivity extends CoreActivity>(
  activity: TActivity,
  app: BotApplication
): ActivityContext<TActivity> {
  return {
    activity,
    send(textOrActivity: string | Partial<CoreActivity>): Promise<ResourceResponse | undefined> {
      const outgoing: Partial<CoreActivity> =
        typeof textOrActivity === 'string'
          ? { type: 'message', text: textOrActivity }
          : textOrActivity;
      return app.sendActivityAsync(
        activity.serviceUrl,
        activity.conversation.id,
        outgoing
      );
    },
  };
}
