/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { Activity } from '@microsoft/agents-activity'

/**
 * Settings for connecting to a Copilot Studio agent via DirectLine.
 */
export interface DirectLineSettings {
  /**
   * The Copilot Studio token endpoint URL.
   * Example: https://{env}.environment.api.powerplatform.com/powervirtualagents/botsbyschema/{bot}/directline/token?api-version=2022-03-01-preview
   */
  tokenEndpoint: string
}

/**
 * Response from the Copilot Studio token endpoint.
 */
export interface TokenResponse {
  token: string
  conversationId?: string
}

/**
 * Response from the regional channel settings endpoint.
 */
export interface RegionalChannelSettings {
  channelUrlsById?: {
    directline?: string
  }
}

/**
 * Represents a started DirectLine conversation.
 */
export interface DirectLineConversation {
  conversationId: string
  token: string
  /** WebSocket URL for streaming activities. */
  streamUrl: string
  /** Last known watermark — can be used to resume after reconnect. */
  watermark?: string
}

/**
 * A set of activities returned by the polling endpoint.
 */
export interface ActivitySet {
  activities: Activity[]
  watermark: string
}

/**
 * Options for the polling listener.
 */
export interface PollingOptions {
  /** Polling interval in milliseconds. Default: 1000. */
  interval?: number
  /** Resume from this watermark — skips activities already seen. */
  watermark?: string
}

/**
 * Options for the WebSocket listener.
 */
export interface WebSocketOptions {
  /** Timeout in ms to wait for the WebSocket to open. Default: 10000. */
  connectTimeout?: number
}

/**
 * Callback for intercepting activities as they arrive from the agent.
 * Return `false` to suppress the activity from the listener's output.
 */
export type ActivityInterceptor = (activity: Activity) => boolean | void

// ---------------------------------------------------------------------------
// Connection Status
// ---------------------------------------------------------------------------

/**
 * Connection status states for the DirectLine listener.
 *
 * State transitions:
 *   Connecting → Connected → Disconnected
 *                Connected → TokenExpired
 *                Connected → Reconnecting → Connected
 *                Connected → Reconnecting → Disconnected
 */
export enum ConnectionStatus {
  /** Establishing the initial connection. */
  Connecting = 'connecting',
  /** Connected and receiving activities. */
  Connected = 'connected',
  /** Retrying after a transient error (polling only). */
  Reconnecting = 'reconnecting',
  /** Connection closed normally or via AbortSignal. */
  Disconnected = 'disconnected',
  /** Token expired (401) — caller must refresh and reconnect. */
  TokenExpired = 'tokenExpired',
}

/**
 * Callback invoked when the listener connection status changes.
 */
export type ConnectionStatusCallback = (status: ConnectionStatus) => void

/**
 * Options shared by both listener methods.
 */
export interface ListenerOptions {
  /** Called when the connection status changes. */
  onStatusChange?: ConnectionStatusCallback
  /** Optional callback to inspect/filter each activity. */
  interceptor?: ActivityInterceptor
}

/**
 * Full options for the polling listener.
 */
export interface PollingListenerOptions extends ListenerOptions, PollingOptions {}

/**
 * Full options for the WebSocket listener.
 */
export interface WebSocketListenerOptions extends ListenerOptions, WebSocketOptions {}
