/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import WebSocket from 'ws'
import { Activity } from '@microsoft/agents-activity'
import {
  DirectLineSettings,
  TokenResponse,
  DirectLineConversation,
  ActivitySet,
  PollingOptions,
  WebSocketOptions,
  ActivityInterceptor,
} from './types'

const DEFAULT_DIRECTLINE_DOMAIN = 'https://directline.botframework.com/v3/directline'

/**
 * DirectLine v3 client for Copilot Studio agents.
 *
 * Handles token acquisition from the CPS token endpoint, regional domain
 * discovery, conversation lifecycle, sending activities, and listening for
 * responses via both HTTP polling and WebSocket modes.
 *
 * Both listener methods return `AsyncGenerator<Activity>` so the consumer
 * pattern is identical regardless of transport.
 *
 * @example
 * ```ts
 * const client = new DirectLineClient({ tokenEndpoint: 'https://...directline/token?api-version=...' })
 * const conversation = await client.startConversation()
 *
 * // WebSocket listener (preferred — real-time, lower latency)
 * for await (const activity of client.listenWebSocket(conversation)) {
 *   console.log(activity.type, activity.text)
 * }
 *
 * // Or polling listener (fallback)
 * for await (const activity of client.listenPolling(conversation)) {
 *   console.log(activity.type, activity.text)
 * }
 * ```
 */
export class DirectLineClient {
  private readonly settings: DirectLineSettings
  private domain: string | undefined
  private token: string | undefined

  constructor (settings: DirectLineSettings) {
    if (!settings.tokenEndpoint) {
      throw new Error('tokenEndpoint is required')
    }
    this.settings = settings
  }

  // ---------------------------------------------------------------------------
  // Token & Domain Discovery
  // ---------------------------------------------------------------------------

  /**
   * Fetches a DirectLine token from the Copilot Studio token endpoint.
   */
  async getToken (): Promise<string> {
    const response = await fetch(this.settings.tokenEndpoint)
    if (!response.ok) {
      throw new Error(`Token fetch failed: ${response.status} ${response.statusText}`)
    }
    const data: TokenResponse = await response.json()
    if (!data.token) {
      throw new Error('Token endpoint response missing "token" field')
    }
    this.token = data.token
    return data.token
  }

  /**
   * Derives the regional channel settings URL from the token endpoint and
   * fetches the regional DirectLine domain.
   *
   * Token URL:    https://{env}.../powervirtualagents/botsbyschema/{bot}/directline/token?api-version={ver}
   * Settings URL: https://{env}.../powervirtualagents/regionalchannelsettings?api-version={ver}
   *
   * Returns the full domain (e.g. "https://europe.directline.botframework.com/v3/directline")
   * or falls back to the global default.
   */
  async discoverDomain (): Promise<string> {
    if (this.domain) return this.domain

    const settingsUrl = getRegionalChannelSettingsUrl(this.settings.tokenEndpoint)
    if (settingsUrl) {
      try {
        const response = await fetch(settingsUrl)
        if (response.ok) {
          const data = await response.json()
          const directLineUrl: string | undefined = data?.channelUrlsById?.directline
          if (directLineUrl) {
            this.domain = directLineUrl.replace(/\/$/, '') + '/v3/directline'
            return this.domain
          }
        }
      } catch {
        // Fall through to default
      }
    }

    this.domain = DEFAULT_DIRECTLINE_DOMAIN
    return this.domain
  }

  // ---------------------------------------------------------------------------
  // Conversation Lifecycle
  // ---------------------------------------------------------------------------

  /**
   * Starts a new DirectLine conversation.
   *
   * Automatically fetches a token and discovers the regional domain if not
   * already done.
   */
  async startConversation (): Promise<DirectLineConversation> {
    const [token, domain] = await Promise.all([
      this.token ? Promise.resolve(this.token) : this.getToken(),
      this.discoverDomain(),
    ])

    const response = await fetch(`${domain}/conversations`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    })

    if (!response.ok) {
      throw new Error(`Start conversation failed: ${response.status} ${response.statusText}`)
    }

    const data = await response.json()
    if (!data.conversationId) {
      throw new Error('Start conversation response missing "conversationId"')
    }

    // The response may include a refreshed token
    if (data.token) {
      this.token = data.token
    }

    return {
      conversationId: data.conversationId,
      token: data.token ?? token,
      streamUrl: data.streamUrl ?? '',
    }
  }

  // ---------------------------------------------------------------------------
  // Send
  // ---------------------------------------------------------------------------

  /**
   * Posts an activity to the conversation.
   * @returns The server-assigned activity ID.
   */
  async sendActivity (conversationId: string, activity: Partial<Activity>): Promise<string> {
    const domain = await this.discoverDomain()
    const token = this.token
    if (!token) throw new Error('No token available. Call getToken() or startConversation() first.')

    const response = await fetch(`${domain}/conversations/${conversationId}/activities`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(activity),
    })

    if (!response.ok) {
      throw new Error(`Send activity failed: ${response.status} ${response.statusText}`)
    }

    const data = await response.json()
    return data.id ?? ''
  }

  // ---------------------------------------------------------------------------
  // Receive — Polling
  // ---------------------------------------------------------------------------

  /**
   * Fetches activities from the conversation since the given watermark.
   */
  async getActivities (conversationId: string, watermark?: string): Promise<ActivitySet> {
    const domain = await this.discoverDomain()
    const token = this.token
    if (!token) throw new Error('No token available.')

    const url = watermark
      ? `${domain}/conversations/${conversationId}/activities?watermark=${watermark}`
      : `${domain}/conversations/${conversationId}/activities`

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    })

    if (!response.ok) {
      throw new Error(`Get activities failed: ${response.status} ${response.statusText}`)
    }

    const data = await response.json()
    return {
      activities: (data.activities ?? []).map((a: any) => Activity.fromObject(a)),
      watermark: data.watermark ?? '',
    }
  }

  /**
   * Polling listener — yields new activities at a configurable interval.
   *
   * Tracks the watermark internally so each iteration only returns unseen
   * activities. The generator runs until the caller breaks out of the loop
   * or the AbortSignal fires.
   *
   * @param conversation - The conversation to listen on.
   * @param options - Polling interval and other options.
   * @param signal - Optional AbortSignal to stop the listener.
   * @param interceptor - Optional callback to inspect/filter each activity.
   */
  async * listenPolling (
    conversation: DirectLineConversation,
    options?: PollingOptions,
    signal?: AbortSignal,
    interceptor?: ActivityInterceptor,
  ): AsyncGenerator<Activity> {
    const interval = options?.interval ?? 1000
    let watermark: string | undefined

    while (!signal?.aborted) {
      const result = await this.getActivities(conversation.conversationId, watermark)
      if (result.watermark) {
        watermark = result.watermark
      }

      for (const activity of result.activities) {
        if (interceptor) {
          const keep = interceptor(activity)
          if (keep === false) continue
        }
        yield activity
      }

      await delay(interval)
    }
  }

  // ---------------------------------------------------------------------------
  // Receive — WebSocket
  // ---------------------------------------------------------------------------

  /**
   * WebSocket listener — yields activities in real-time as they arrive on
   * the WebSocket stream.
   *
   * The `streamUrl` from `startConversation()` is a one-time-use URL.
   * When the server closes the connection it provides a new streamUrl in the
   * last message; this listener does NOT auto-reconnect — the caller should
   * start a new listener if needed.
   *
   * @param conversation - Must include a valid `streamUrl`.
   * @param options - Connection timeout and other options.
   * @param signal - Optional AbortSignal to close the socket.
   * @param interceptor - Optional callback to inspect/filter each activity.
   */
  async * listenWebSocket (
    conversation: DirectLineConversation,
    options?: WebSocketOptions,
    signal?: AbortSignal,
    interceptor?: ActivityInterceptor,
  ): AsyncGenerator<Activity> {
    if (!conversation.streamUrl) {
      throw new Error('No streamUrl available. Start a conversation first.')
    }

    const connectTimeout = options?.connectTimeout ?? 10_000

    const ws = new WebSocket(conversation.streamUrl)
    const messageQueue: Activity[] = []
    let resolve: (() => void) | undefined
    let error: Error | undefined
    let closed = false

    ws.on('message', (data: WebSocket.Data) => {
      try {
        const parsed = JSON.parse(data.toString())
        const activities: any[] = parsed.activities ?? []
        for (const raw of activities) {
          const activity = Activity.fromObject(raw)
          if (interceptor) {
            const keep = interceptor(activity)
            if (keep === false) continue
          }
          messageQueue.push(activity)
        }
        // Wake up the consumer if it's waiting
        resolve?.()
      } catch {
        // Ignore unparseable frames
      }
    })

    ws.on('close', () => {
      closed = true
      resolve?.()
    })

    ws.on('error', (err: Error) => {
      error = err
      closed = true
      resolve?.()
    })

    // Handle abort signal
    const onAbort = () => {
      ws.close()
    }
    signal?.addEventListener('abort', onAbort, { once: true })

    // Wait for connection to open
    await new Promise<void>((res, rej) => {
      const timeout = setTimeout(() => {
        ws.close()
        rej(new Error(`WebSocket connection timed out after ${connectTimeout}ms`))
      }, connectTimeout)

      ws.on('open', () => {
        clearTimeout(timeout)
        res()
      })

      ws.on('error', (err) => {
        clearTimeout(timeout)
        rej(err)
      })
    })

    try {
      while (!closed && !signal?.aborted) {
        // Drain the queue
        while (messageQueue.length > 0) {
          yield messageQueue.shift()!
        }

        if (closed || signal?.aborted) break

        // Wait for the next message or close
        await new Promise<void>((res) => {
          resolve = res
        })

        if (error) {
          throw error
        }
      }

      // Drain any remaining messages
      while (messageQueue.length > 0) {
        yield messageQueue.shift()!
      }
    } finally {
      signal?.removeEventListener('abort', onAbort)
      if (ws.readyState === WebSocket.OPEN || ws.readyState === WebSocket.CONNECTING) {
        ws.close()
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Derives the regional channel settings URL from a CPS token endpoint.
 */
function getRegionalChannelSettingsUrl (tokenEndpoint: string): string | null {
  const pvaIndex = tokenEndpoint.indexOf('/powervirtualagents')
  if (pvaIndex === -1) return null

  const environmentEndpoint = tokenEndpoint.slice(0, pvaIndex)
  const apiVersionMatch = tokenEndpoint.match(/api-version=([^&]+)/)
  const apiVersion = apiVersionMatch ? apiVersionMatch[1] : '2022-03-01-preview'

  return `${environmentEndpoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`
}

function delay (ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms))
}
