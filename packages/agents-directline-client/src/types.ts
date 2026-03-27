/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { Activity } from '@microsoft/agents-activity'

// ---------------------------------------------------------------------------
// Public API types
// ---------------------------------------------------------------------------

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
 *
 * Note: If the initial connection fails (before ever reaching Connected),
 * the status remains Connecting until Disconnected — Reconnecting is only
 * emitted after a prior successful connection.
 */
export enum ConnectionStatus {
  /** Establishing the initial connection. */
  Connecting = 'connecting',
  /** Connected and receiving activities. */
  Connected = 'connected',
  /** Retrying after a transient error (only after a prior Connected state). */
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

// ---------------------------------------------------------------------------
// Listener Options
// ---------------------------------------------------------------------------

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
export interface PollingListenerOptions extends ListenerOptions {
  /** Polling interval in milliseconds. Default: 1000. */
  interval?: number
  /** Resume from this watermark — skips activities already seen. */
  watermark?: string
}

/**
 * Full options for the WebSocket listener.
 */
export interface WebSocketListenerOptions extends ListenerOptions {
  /** Timeout in ms to wait for the WebSocket to open. Default: 10000. */
  connectTimeout?: number
}
