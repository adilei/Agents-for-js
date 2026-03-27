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
