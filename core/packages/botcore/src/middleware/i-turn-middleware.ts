// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { CoreActivity } from '../schema/core-activity.js'

/** Callback to invoke the next middleware (or the activity handler) in the pipeline. */
export type NextTurn = () => Promise<void>

/**
 * A middleware that participates in the turn processing pipeline.
 *
 * Implement this interface to add cross-cutting logic (e.g. auth, logging,
 * telemetry) that runs before or after each activity is handled.
 *
 * @example
 * class LoggingMiddleware implements ITurnMiddleware {
 *   async onTurnAsync(app, activity, next) {
 *     console.log('activity type:', activity.type);
 *     await next();
 *   }
 * }
 */
export interface ITurnMiddleware {
  /**
   * Called once per incoming activity turn.
   *
   * @param app - The `BotApplication` instance processing the turn.
   * @param activity - The incoming activity.
   * @param next - Call this to continue to the next middleware or handler.
   */
  onTurnAsync(
    app: unknown,
    activity: CoreActivity,
    next: NextTurn
  ): Promise<void>;
}
