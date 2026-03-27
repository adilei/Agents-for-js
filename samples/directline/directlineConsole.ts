/**
 * DirectLine Console Sample — Two-Thread Pattern
 *
 * Demonstrates decoupled send/receive using the DirectLine v3 client.
 * Two independent async loops run concurrently:
 *   - Input loop:    readline → sendActivity()
 *   - Listener loop: WebSocket (or polling) → print activities
 *
 * This is the foundation for the relay pattern: sending and receiving
 * are fully independent, enabling proactive messaging and background
 * listening.
 *
 * Configuration:
 *   Copy .env.example to .env and set DIRECTLINE_TOKEN_ENDPOINT to your
 *   Copilot Studio token endpoint URL.
 *
 * Usage:
 *   npx tsx --env-file=samples/directline/.env samples/directline/directlineConsole.ts
 *   npx tsx --env-file=samples/directline/.env samples/directline/directlineConsole.ts --poll
 */

import readline from 'readline'
import { Activity, ActivityTypes, CardAction } from '@microsoft/agents-activity'
import { DirectLineClient, DirectLineConversation, DirectLineSettings } from '@microsoft/agents-directline-client'

// ---------------------------------------------------------------------------
// Settings (loaded from environment / .env file)
// ---------------------------------------------------------------------------

function loadSettings (): DirectLineSettings {
  const tokenEndpoint = process.env.DIRECTLINE_TOKEN_ENDPOINT
  if (!tokenEndpoint) {
    console.error('Error: DIRECTLINE_TOKEN_ENDPOINT is required.')
    console.error('Copy .env.example to .env and set the token endpoint URL.')
    console.error('Example: https://{env}.environment.api.powerplatform.com/powervirtualagents/botsbyschema/{bot}/directline/token?api-version=2022-03-01-preview')
    process.exit(1)
  }
  return { tokenEndpoint }
}

function getTransportMode (): 'websocket' | 'polling' {
  if (process.argv.includes('--poll')) return 'polling'
  const envTransport = process.env.DIRECTLINE_TRANSPORT?.toLowerCase()
  if (envTransport === 'polling') return 'polling'
  return 'websocket'
}

const settings = loadSettings()
const transport = getTransportMode()

// ---------------------------------------------------------------------------
// Activity Printer
// ---------------------------------------------------------------------------

function printActivity (activity: Activity): void {
  switch (activity.type) {
    case ActivityTypes.Message:
      if (activity.text) {
        console.log(`\n  🤖 ${activity.text}`)
      }
      activity.suggestedActions?.actions?.forEach((action: CardAction) =>
        console.log(`     → ${action.title ?? action.value}`)
      )
      if (activity.attachments?.length) {
        for (const att of activity.attachments) {
          console.log(`     📎 [${att.contentType}] ${att.name ?? att.contentUrl ?? ''}`)
        }
      }
      break
    case ActivityTypes.Typing:
      process.stdout.write('  ...')
      break
    case ActivityTypes.EndOfConversation:
      console.log('\n  --- Conversation ended ---')
      break
    default:
      console.log(`\n  (${activity.type}) ${activity.text ?? ''}`)
  }
}

// ---------------------------------------------------------------------------
// Listener Loop (background)
// ---------------------------------------------------------------------------

async function listenInBackground (
  client: DirectLineClient,
  conversation: DirectLineConversation,
  signal: AbortSignal,
): Promise<void> {
  const mode = transport === 'polling' ? 'polling' : 'WebSocket'
  console.log(`\n  [listener] Started (${mode} mode)\n`)

  try {
    const listener = transport === 'polling'
      ? client.listenPolling(conversation, { interval: 1000 }, signal)
      : client.listenWebSocket(conversation, {}, signal)

    for await (const activity of listener) {
      printActivity(activity)
      // Re-show the prompt after printing
      process.stdout.write('\n>>>: ')
    }
  } catch (err: any) {
    if (!signal.aborted) {
      console.error(`\n  [listener] Error: ${err.message}`)
    }
  }

  console.log('  [listener] Stopped')
}

// ---------------------------------------------------------------------------
// Input Loop (foreground)
// ---------------------------------------------------------------------------

async function inputLoop (
  client: DirectLineClient,
  conversationId: string,
  abort: AbortController,
): Promise<void> {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  })

  const ask = (): Promise<string> =>
    new Promise((resolve) => rl.question('>>>: ', resolve))

  try {
    while (!abort.signal.aborted) {
      const input = await ask()

      if (input.toLowerCase() === 'exit') {
        console.log('\n  Shutting down...')
        break
      }

      if (!input.trim()) continue

      try {
        await client.sendActivity(conversationId, {
          type: 'message',
          text: input,
          from: { id: 'user', name: 'User' },
        })
      } catch (err: any) {
        console.error(`  Send error: ${err.message}`)
      }
    }
  } finally {
    rl.close()
    abort.abort()
  }
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main (): Promise<void> {
  console.log('DirectLine Console Sample')
  console.log('='.repeat(40))
  console.log(`  Mode: ${transport === 'polling' ? 'HTTP Polling' : 'WebSocket'}`)
  console.log(`  Token endpoint: ${settings.tokenEndpoint.substring(0, 60)}...`)
  console.log()

  const client = new DirectLineClient(settings)

  console.log('  Connecting...')
  const conversation = await client.startConversation()
  console.log(`  Connected! Conversation ID: ${conversation.conversationId}`)
  console.log(`  Type a message and press Enter. Type "exit" to quit.\n`)

  const abort = new AbortController()

  // Launch listener in the background (not awaited)
  const listenerPromise = listenInBackground(client, conversation, abort.signal)

  // Run input loop in the foreground (blocks until exit)
  await inputLoop(client, conversation.conversationId, abort)

  // Wait for listener to finish after abort
  await listenerPromise

  console.log('  Goodbye!')
  process.exit(0)
}

main().catch((err) => {
  console.error('Fatal error:', err.message)
  process.exit(1)
})
