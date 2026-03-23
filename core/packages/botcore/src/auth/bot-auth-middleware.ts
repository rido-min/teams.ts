// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as jwt from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import axios from 'axios';
import createDebug from 'debug';

const debug = createDebug('teams:botcore:auth');

const BOT_OPENID_METADATA_URL =
  'https://login.botframework.com/v1/.well-known/openidconfiguration';

let cachedJwksUri: string | undefined;
let cachedJwks: ReturnType<typeof jwksClient> | undefined;

async function getJwks(): Promise<ReturnType<typeof jwksClient>> {
  if (!cachedJwks) {
    debug('fetching JWKS URI from %s', BOT_OPENID_METADATA_URL);
    const resp = await axios.get<{ jwks_uri: string }>(BOT_OPENID_METADATA_URL);
    cachedJwksUri = resp.data.jwks_uri;
    debug('JWKS URI resolved: %s', cachedJwksUri);
    cachedJwks = jwksClient({ jwksUri: cachedJwksUri });
  }
  return cachedJwks;
}

/**
 * Validates a Bot Framework Bearer token.
 * Throws if the token is missing, malformed, or fails signature/audience checks.
 *
 * @param authHeader - The Authorization header value (e.g. "Bearer <token>")
 * @param appId - The bot's client ID (audience). Falls back to CLIENT_ID env var.
 */
export async function validateBotToken(
  authHeader: string | undefined,
  appId?: string
): Promise<void> {
  const audience = appId ?? process.env['CLIENT_ID'];
  if (!audience) throw new BotAuthError('No appId configured for token validation');

  if (!authHeader?.startsWith('Bearer ')) {
    debug('auth failed: missing or malformed Authorization header');
    throw new BotAuthError('Missing or invalid Authorization header');
  }

  const token = authHeader.slice(7);
  debug('validating token for audience %s', audience);

  const jwks = await getJwks();

  await new Promise<void>((resolve, reject) => {
    jwt.verify(
      token,
      (header, callback) => {
        debug('fetching signing key kid=%s', header.kid);
        jwks.getSigningKey(header.kid, (err, key) => {
          if (err) {
            debug('signing key fetch failed: %s', err.message);
            return callback(err, undefined);
          }
          callback(null, key?.getPublicKey());
        });
      },
      { audience, algorithms: ['RS256'] },
      (err) => {
        if (err) {
          debug('token validation failed: %s', err.message);
          reject(new BotAuthError(`Token validation failed: ${err.message}`));
        } else {
          debug('token validated successfully');
          resolve();
        }
      }
    );
  });
}

export class BotAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'BotAuthError';
  }
}

// ── Express middleware ────────────────────────────────────────────────────────

type ExpressRequest = { headers: Record<string, string | string[] | undefined> };
type ExpressResponse = { status(code: number): ExpressResponse; end(msg?: string): void };
type NextFn = (err?: unknown) => void;

/**
 * Express middleware that validates the Bot Framework JWT token.
 *
 * @example
 * app.post('/api/messages', botAuthExpress(appId), (req, res) => bot.processAsync(req, res));
 */
export function botAuthExpress(
  appId?: string
): (req: ExpressRequest, res: ExpressResponse, next: NextFn) => Promise<void> {
  return async (req, res, next) => {
    const header = req.headers['authorization'] as string | undefined;
    try {
      await validateBotToken(header, appId);
      next();
    } catch (err: unknown) {
      if (err instanceof BotAuthError) {
        res.status(401).end(err.message);
      } else {
        next(err);
      }
    }
  };
}

// ── Hono middleware ───────────────────────────────────────────────────────────

type HonoContext = {
  req: { header(name: string): string | undefined };
  text(body: string, status?: number): Response;
};
type HonoNext = () => Promise<Response | void>;

/**
 * Hono middleware that validates the Bot Framework JWT token.
 *
 * @example
 * honoApp.post('/api/messages', botAuthHono(appId), handler);
 */
export function botAuthHono(
  appId?: string
): (c: HonoContext, next: HonoNext) => Promise<Response | void> {
  return async (c, next) => {
    const header = c.req.header('authorization');
    try {
      await validateBotToken(header, appId);
      return next();
    } catch (err) {
      if (err instanceof BotAuthError) {
        return c.text(err.message, 401);
      }
      throw err;
    }
  };
}
