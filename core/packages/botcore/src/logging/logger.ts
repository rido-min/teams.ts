// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import createDebug from 'debug'

/**
 * Minimal logger interface. Implement this to integrate with any logging
 * framework (pino, winston, console, etc.).
 *
 * Log levels follow standard severity order:
 * `trace` < `debug` < `info` < `warn` < `error`
 */
export interface Logger {
  /** Verbose diagnostic detail — API URLs, full payloads. */
  trace(message: string, ...args: unknown[]): void
  /** Configuration and internal flow information. */
  debug(message: string, ...args: unknown[]): void
  /** Key operational events — activity received, token acquired. */
  info(message: string, ...args: unknown[]): void
  /** Recoverable issues — running without auth, unexpected responses. */
  warn(message: string, ...args: unknown[]): void
  /** Failures during activity processing or API calls. */
  error(message: string, ...args: unknown[]): void
}

/** A logger that discards all messages. */
export const noopLogger: Logger = {
  trace () {},
  debug () {},
  info () {},
  warn () {},
  error () {},
}

/**
 * Default logger backed by the `debug` package.
 *
 * Each level maps to a separate namespace so you can filter precisely:
 *
 * ```sh
 * DEBUG=teams:botcore:*            # all levels
 * DEBUG=teams:botcore:info,teams:botcore:warn,teams:botcore:error  # info and above
 * DEBUG=teams:botcore:error        # errors only
 * ```
 */
export const debugLogger: Logger = {
  trace: createDebug('teams:botcore:trace'),
  debug: createDebug('teams:botcore:debug'),
  info: createDebug('teams:botcore:info'),
  warn: createDebug('teams:botcore:warn'),
  error: createDebug('teams:botcore:error'),
}

/**
 * A logger that writes to `console` with a level prefix.
 *
 * @example
 * import { configure, consoleLogger } from '@microsoft/teams.botcore'
 * configure(consoleLogger)
 */
export const consoleLogger: Logger = {
  trace (message, ...args) { console.debug(`[TRACE] ${message}`, ...args) },
  debug (message, ...args) { console.debug(`[DEBUG] ${message}`, ...args) },
  info (message, ...args) { console.info(`[INFO]  ${message}`, ...args) },
  warn (message, ...args) { console.warn(`[WARN]  ${message}`, ...args) },
  error (message, ...args) { console.error(`[ERROR] ${message}`, ...args) },
}

let _logger: Logger = debugLogger

/**
 * Configure the logger used by all `@microsoft/teams.botcore` components.
 * Call this once at application startup, before creating a `BotApplication`.
 *
 * @example
 * import { configure, consoleLogger } from '@microsoft/teams.botcore'
 * configure(consoleLogger)
 */
export function configure (logger: Logger): void {
  _logger = logger
}

/** @internal */
export function getLogger (): Logger {
  return _logger
}
