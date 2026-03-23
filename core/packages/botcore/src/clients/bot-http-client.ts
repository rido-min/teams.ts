// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import axios, { type AxiosInstance, type AxiosRequestConfig } from 'axios'

/** Options for a single HTTP request made by {@link BotHttpClient}. */
export interface BotRequestOptions {
  /** Human-readable description of the operation (used in error messages). */
  operationDescription?: string;
  /** When `true`, a 404 response returns `undefined` instead of throwing. */
  returnNullOnNotFound?: boolean;
  /** Additional request headers to merge in. */
  headers?: Record<string, string>;
}

/** A function that returns a Bearer token for the outgoing request. */
export type TokenProvider = () => Promise<string>

/**
 * Authenticated HTTP client for the Bot Framework REST API.
 *
 * Wraps axios and injects a Bearer token via a request interceptor when a
 * {@link TokenProvider} is supplied.
 */
export class BotHttpClient {
  private readonly http: AxiosInstance

  /**
   * @param getToken - Optional token provider. When supplied, every request
   *   receives an `Authorization: Bearer <token>` header.
   */
  constructor (getToken?: TokenProvider) {
    this.http = axios.create({
      headers: { 'Content-Type': 'application/json' },
    })

    if (getToken) {
      this.http.interceptors.request.use(async (config) => {
        const token = await getToken()
        config.headers.Authorization = `Bearer ${token}`
        return config
      })
    }
  }

  /**
   * Send an HTTP request and return the parsed response body.
   *
   * @param method - HTTP method (e.g. `"GET"`, `"POST"`).
   * @param url - Fully-qualified URL.
   * @param body - Optional request body.
   * @param options - Request options.
   * @returns Parsed response body, or `undefined` when
   *   `returnNullOnNotFound` is set and the server returns 404.
   */
  async send<T>(
    method: string,
    url: string,
    body?: unknown,
    options: BotRequestOptions = {}
  ): Promise<T | undefined> {
    const config: AxiosRequestConfig = {
      method,
      url,
      data: body ?? undefined,
      headers: options.headers,
    }

    try {
      const response = await this.http.request<T>(config)
      return response.data ?? undefined
    } catch (err: unknown) {
      if (
        options.returnNullOnNotFound &&
        isAxiosError(err) &&
        err.response?.status === 404
      ) {
        return undefined
      }
      throw err
    }
  }

  /**
   * Send a GET request.
   *
   * @param baseUrl - Base service URL (scheme + host).
   * @param endpoint - Path relative to `baseUrl`.
   * @param params - Optional query-string parameters (undefined values omitted).
   * @param options - Request options.
   */
  async get<T>(
    baseUrl: string,
    endpoint: string,
    params?: Record<string, string | undefined>,
    options?: BotRequestOptions
  ): Promise<T | undefined> {
    const url = buildUrl(baseUrl, endpoint, params)
    return this.send<T>('GET', url, undefined, options)
  }

  /**
   * Send a POST request.
   *
   * @param baseUrl - Base service URL (scheme + host).
   * @param endpoint - Path relative to `baseUrl`.
   * @param body - Request body.
   * @param options - Request options.
   */
  async post<T>(
    baseUrl: string,
    endpoint: string,
    body?: unknown,
    options?: BotRequestOptions
  ): Promise<T | undefined> {
    const url = buildUrl(baseUrl, endpoint)
    return this.send<T>('POST', url, body, options)
  }

  /**
   * Send a PUT request.
   *
   * @param baseUrl - Base service URL (scheme + host).
   * @param endpoint - Path relative to `baseUrl`.
   * @param body - Request body.
   * @param options - Request options.
   */
  async put<T>(
    baseUrl: string,
    endpoint: string,
    body?: unknown,
    options?: BotRequestOptions
  ): Promise<T | undefined> {
    const url = buildUrl(baseUrl, endpoint)
    return this.send<T>('PUT', url, body, options)
  }

  /**
   * Send a DELETE request.
   *
   * @param baseUrl - Base service URL (scheme + host).
   * @param endpoint - Path relative to `baseUrl`.
   * @param params - Optional query-string parameters (undefined values omitted).
   * @param options - Request options.
   */
  async delete (
    baseUrl: string,
    endpoint: string,
    params?: Record<string, string | undefined>,
    options?: BotRequestOptions
  ): Promise<void> {
    const url = buildUrl(baseUrl, endpoint, params)
    await this.send<void>('DELETE', url, undefined, options)
  }
}

function buildUrl (
  baseUrl: string,
  endpoint: string,
  params?: Record<string, string | undefined>
): string {
  const base = baseUrl.replace(/\/$/, '')
  const path = endpoint.startsWith('/') ? endpoint : `/${endpoint}`
  let url = `${base}${path}`

  if (params) {
    const query = Object.entries(params)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v!)}`)
      .join('&')
    if (query) url += `?${query}`
  }

  return url
}

function isAxiosError (err: unknown): err is { response?: { status: number } } {
  return typeof err === 'object' && err !== null && 'response' in err
}
