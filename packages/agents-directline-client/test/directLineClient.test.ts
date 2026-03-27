import { strict as assert } from 'assert'
import { describe, it, mock, beforeEach, afterEach } from 'node:test'
import { DirectLineClient } from '../src/directLineClient'
import { Activity } from '@microsoft/agents-activity'
import WebSocket, { WebSocketServer } from 'ws'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Create a mock fetch that returns canned responses based on URL patterns.
 *  Patterns are matched longest-first so "activities" matches before "conversations". */
function createMockFetch (handlers: Record<string, () => { status: number, body: any }>) {
  const sortedKeys = Object.keys(handlers).sort((a, b) => b.length - a.length)
  return mock.fn(async (url: string | URL | Request, init?: RequestInit) => {
    const urlStr = typeof url === 'string' ? url : url instanceof URL ? url.href : url.url
    for (const pattern of sortedKeys) {
      if (urlStr.includes(pattern)) {
        const { status, body } = handlers[pattern]()
        return new Response(JSON.stringify(body), {
          status,
          headers: { 'Content-Type': 'application/json' },
        })
      }
    }
    return new Response('Not found', { status: 404 })
  })
}

const TOKEN_ENDPOINT = 'https://abc123.42.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cr26e_myBot/directline/token?api-version=2022-03-01-preview'

// ---------------------------------------------------------------------------
// Regional Domain Discovery
// ---------------------------------------------------------------------------

describe('DirectLineClient', () => {
  let originalFetch: typeof globalThis.fetch

  beforeEach(() => {
    originalFetch = globalThis.fetch
  })

  afterEach(() => {
    globalThis.fetch = originalFetch
  })

  describe('constructor', () => {
    it('throws when tokenEndpoint is empty', () => {
      assert.throws(() => new DirectLineClient({ tokenEndpoint: '' }), /tokenEndpoint is required/)
    })

    it('creates instance with valid settings', () => {
      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      assert.ok(client)
    })
  })

  describe('getToken', () => {
    it('fetches token from the token endpoint', async () => {
      globalThis.fetch = createMockFetch({
        'directline/token': () => ({ status: 200, body: { token: 'test-token-123' } }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const token = await client.getToken()

      assert.equal(token, 'test-token-123')
    })

    it('throws when token endpoint returns error', async () => {
      globalThis.fetch = createMockFetch({
        'directline/token': () => ({ status: 401, body: { error: 'unauthorized' } }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      await assert.rejects(client.getToken(), /Token fetch failed: 401/)
    })

    it('throws when response is missing token field', async () => {
      globalThis.fetch = createMockFetch({
        'directline/token': () => ({ status: 200, body: { conversationId: 'abc' } }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      await assert.rejects(client.getToken(), /missing "token" field/)
    })
  })

  describe('discoverDomain', () => {
    it('discovers regional domain from channel settings', async () => {
      globalThis.fetch = createMockFetch({
        'regionalchannelsettings': () => ({
          status: 200,
          body: {
            channelUrlsById: {
              directline: 'https://europe.directline.botframework.com',
            },
          },
        }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const domain = await client.discoverDomain()

      assert.equal(domain, 'https://europe.directline.botframework.com/v3/directline')
    })

    it('falls back to global default when regional settings fail', async () => {
      globalThis.fetch = createMockFetch({
        'regionalchannelsettings': () => ({ status: 500, body: {} }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const domain = await client.discoverDomain()

      assert.equal(domain, 'https://directline.botframework.com/v3/directline')
    })

    it('falls back when token endpoint has no /powervirtualagents path', async () => {
      const client = new DirectLineClient({
        tokenEndpoint: 'https://example.com/some/other/token/endpoint',
      })
      // No fetch mock needed — it should not call fetch at all for settings
      globalThis.fetch = mock.fn(async () => new Response('', { status: 404 })) as any

      const domain = await client.discoverDomain()
      assert.equal(domain, 'https://directline.botframework.com/v3/directline')
    })

    it('caches the discovered domain', async () => {
      const fetchMock = createMockFetch({
        'regionalchannelsettings': () => ({
          status: 200,
          body: { channelUrlsById: { directline: 'https://asia.directline.botframework.com' } },
        }),
      })
      globalThis.fetch = fetchMock as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const domain1 = await client.discoverDomain()
      const domain2 = await client.discoverDomain()

      assert.equal(domain1, domain2)
      assert.equal(fetchMock.mock.callCount(), 1)
    })
  })

  describe('startConversation', () => {
    it('starts a conversation and returns conversationId + streamUrl', async () => {
      globalThis.fetch = createMockFetch({
        'directline/token': () => ({
          status: 200,
          body: { token: 'tok-1' },
        }),
        'regionalchannelsettings': () => ({
          status: 200,
          body: { channelUrlsById: { directline: 'https://directline.botframework.com' } },
        }),
        'conversations': () => ({
          status: 200,
          body: {
            conversationId: 'conv-abc',
            token: 'tok-refreshed',
            streamUrl: 'wss://directline.botframework.com/v3/directline/conversations/conv-abc/stream?t=tok-refreshed',
          },
        }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const conv = await client.startConversation()

      assert.equal(conv.conversationId, 'conv-abc')
      assert.equal(conv.token, 'tok-refreshed')
      assert.ok(conv.streamUrl.includes('conv-abc'))
    })
  })

  describe('sendActivity', () => {
    it('posts an activity and returns the activity id', async () => {
      globalThis.fetch = createMockFetch({
        'directline/token': () => ({ status: 200, body: { token: 'tok-1' } }),
        'regionalchannelsettings': () => ({
          status: 200,
          body: { channelUrlsById: { directline: 'https://directline.botframework.com' } },
        }),
        'conv-1/activities': () => ({
          status: 200,
          body: { id: 'act-001' },
        }),
        'conversations': () => ({
          status: 200,
          body: { conversationId: 'conv-1', token: 'tok-1', streamUrl: '' },
        }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      await client.startConversation()
      const actId = await client.sendActivity('conv-1', { type: 'message', text: 'hello' })

      assert.equal(actId, 'act-001')
    })
  })

  describe('getActivities', () => {
    it('returns activities and watermark', async () => {
      globalThis.fetch = createMockFetch({
        'directline/token': () => ({ status: 200, body: { token: 'tok-1' } }),
        'regionalchannelsettings': () => ({
          status: 200,
          body: { channelUrlsById: { directline: 'https://directline.botframework.com' } },
        }),
        'activities': () => ({
          status: 200,
          body: {
            activities: [
              { type: 'message', text: 'Hello from bot', from: { id: 'bot' } },
            ],
            watermark: '1',
          },
        }),
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      await client.getToken()
      const result = await client.getActivities('conv-1')

      assert.equal(result.activities.length, 1)
      assert.equal(result.activities[0].text, 'Hello from bot')
      assert.equal(result.watermark, '1')
    })

    it('passes watermark as query parameter', async () => {
      let capturedUrl = ''
      globalThis.fetch = mock.fn(async (url: string | URL | Request) => {
        const urlStr = typeof url === 'string' ? url : url instanceof URL ? url.href : url.url
        capturedUrl = urlStr
        if (urlStr.includes('regionalchannelsettings')) {
          return new Response(JSON.stringify({ channelUrlsById: { directline: 'https://directline.botframework.com' } }), { status: 200 })
        }
        if (urlStr.includes('directline/token')) {
          return new Response(JSON.stringify({ token: 'tok-1' }), { status: 200 })
        }
        return new Response(JSON.stringify({ activities: [], watermark: '5' }), { status: 200 })
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      await client.getToken()
      await client.getActivities('conv-1', '3')

      assert.ok(capturedUrl.includes('watermark=3'))
    })
  })

  describe('listenPolling', () => {
    it('yields activities and respects abort signal', async () => {
      let callCount = 0
      globalThis.fetch = mock.fn(async (url: string | URL | Request) => {
        const urlStr = typeof url === 'string' ? url : url instanceof URL ? url.href : url.url
        if (urlStr.includes('regionalchannelsettings')) {
          return new Response(JSON.stringify({ channelUrlsById: { directline: 'https://directline.botframework.com' } }), { status: 200 })
        }
        if (urlStr.includes('directline/token')) {
          return new Response(JSON.stringify({ token: 'tok-1' }), { status: 200 })
        }
        // Return activities on first poll, empty on second
        callCount++
        if (callCount === 1) {
          return new Response(JSON.stringify({
            activities: [
              { type: 'message', text: 'reply 1', from: { id: 'bot' } },
              { type: 'message', text: 'reply 2', from: { id: 'bot' } },
            ],
            watermark: '2',
          }), { status: 200 })
        }
        return new Response(JSON.stringify({ activities: [], watermark: '2' }), { status: 200 })
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      await client.getToken()

      const controller = new AbortController()
      const conversation = { conversationId: 'conv-1', token: 'tok-1', streamUrl: '' }
      const collected: Activity[] = []

      for await (const activity of client.listenPolling(conversation, { interval: 50 }, controller.signal)) {
        collected.push(activity)
        if (collected.length >= 2) {
          controller.abort()
        }
      }

      assert.equal(collected.length, 2)
      assert.equal(collected[0].text, 'reply 1')
      assert.equal(collected[1].text, 'reply 2')
    })

    it('applies interceptor to filter activities', async () => {
      let callCount = 0
      globalThis.fetch = mock.fn(async (url: string | URL | Request) => {
        const urlStr = typeof url === 'string' ? url : url instanceof URL ? url.href : url.url
        if (urlStr.includes('regionalchannelsettings')) {
          return new Response(JSON.stringify({ channelUrlsById: { directline: 'https://directline.botframework.com' } }), { status: 200 })
        }
        if (urlStr.includes('directline/token')) {
          return new Response(JSON.stringify({ token: 'tok-1' }), { status: 200 })
        }
        callCount++
        if (callCount === 1) {
          return new Response(JSON.stringify({
            activities: [
              { type: 'typing', from: { id: 'bot' } },
              { type: 'message', text: 'real reply', from: { id: 'bot' } },
            ],
            watermark: '2',
          }), { status: 200 })
        }
        return new Response(JSON.stringify({ activities: [], watermark: '2' }), { status: 200 })
      }) as any

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      await client.getToken()

      const controller = new AbortController()
      const conversation = { conversationId: 'conv-1', token: 'tok-1', streamUrl: '' }
      const collected: Activity[] = []

      // Only allow message activities through
      const interceptor = (a: Activity) => a.type === 'message'

      for await (const activity of client.listenPolling(conversation, { interval: 50 }, controller.signal, interceptor)) {
        collected.push(activity)
        controller.abort()
      }

      assert.equal(collected.length, 1)
      assert.equal(collected[0].type, 'message')
      assert.equal(collected[0].text, 'real reply')
    })
  })

  describe('listenWebSocket', () => {
    let wss: WebSocketServer

    afterEach(() => {
      if (wss) {
        wss.close()
      }
    })

    it('yields activities from WebSocket stream', async () => {
      // Start a local WebSocket server
      wss = new WebSocketServer({ port: 0 })
      const port = (wss.address() as any).port

      wss.on('connection', (ws) => {
        // Send an activity set after a short delay
        setTimeout(() => {
          ws.send(JSON.stringify({
            activities: [
              { type: 'message', text: 'ws reply 1', from: { id: 'bot', name: 'Bot' } },
              { type: 'message', text: 'ws reply 2', from: { id: 'bot', name: 'Bot' } },
            ],
          }))
        }, 50)

        // Close after sending
        setTimeout(() => ws.close(), 150)
      })

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const conversation = {
        conversationId: 'conv-ws',
        token: 'tok-1',
        streamUrl: `ws://localhost:${port}`,
      }

      const collected: Activity[] = []
      for await (const activity of client.listenWebSocket(conversation)) {
        collected.push(activity)
      }

      assert.equal(collected.length, 2)
      assert.equal(collected[0].text, 'ws reply 1')
      assert.equal(collected[1].text, 'ws reply 2')
    })

    it('stops on abort signal', async () => {
      wss = new WebSocketServer({ port: 0 })
      const port = (wss.address() as any).port

      wss.on('connection', (ws) => {
        // Send activities periodically
        const interval = setInterval(() => {
          if (ws.readyState === WebSocket.OPEN) {
            ws.send(JSON.stringify({
              activities: [{ type: 'message', text: 'tick', from: { id: 'bot' } }],
            }))
          }
        }, 50)
        ws.on('close', () => clearInterval(interval))
      })

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const conversation = {
        conversationId: 'conv-ws',
        token: 'tok-1',
        streamUrl: `ws://localhost:${port}`,
      }

      const controller = new AbortController()
      const collected: Activity[] = []

      for await (const activity of client.listenWebSocket(conversation, {}, controller.signal)) {
        collected.push(activity)
        if (collected.length >= 2) {
          controller.abort()
        }
      }

      assert.ok(collected.length >= 2)
      assert.equal(collected[0].text, 'tick')
    })

    it('throws when no streamUrl is provided', async () => {
      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const conversation = { conversationId: 'conv-1', token: 'tok-1', streamUrl: '' }

      await assert.rejects(async () => {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        for await (const _ of client.listenWebSocket(conversation)) {
          // Should not reach here
        }
      }, /No streamUrl available/)
    })

    it('applies interceptor to WebSocket activities', async () => {
      wss = new WebSocketServer({ port: 0 })
      const port = (wss.address() as any).port

      wss.on('connection', (ws) => {
        setTimeout(() => {
          ws.send(JSON.stringify({
            activities: [
              { type: 'typing', from: { id: 'bot' } },
              { type: 'message', text: 'kept', from: { id: 'bot' } },
            ],
          }))
        }, 50)
        setTimeout(() => ws.close(), 150)
      })

      const client = new DirectLineClient({ tokenEndpoint: TOKEN_ENDPOINT })
      const conversation = {
        conversationId: 'conv-ws',
        token: 'tok-1',
        streamUrl: `ws://localhost:${port}`,
      }

      const collected: Activity[] = []
      const interceptor = (a: Activity) => a.type !== 'typing'

      for await (const activity of client.listenWebSocket(conversation, {}, undefined, interceptor)) {
        collected.push(activity)
      }

      assert.equal(collected.length, 1)
      assert.equal(collected[0].text, 'kept')
    })
  })
})
