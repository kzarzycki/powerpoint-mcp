/* PowerPoint MCP — WebSocket client */

var ws = null
var reconnectAttempt = 0
var BASE_DELAY = 500
var MAX_DELAY = 30000
var officeReady = false
var AsyncFunction = (async () => {}).constructor

/* Office.js initialization */
Office.onReady((info) => {
  officeReady = true
  console.log('Office.js ready:', info.host, info.platform)
  updateStatus('connecting')
  initWebSocket()
})

/* Fallback for browser testing (no Office.js host) */
setTimeout(() => {
  if (!officeReady) {
    console.log('Office.js not detected — standalone mode')
    updateStatus('connecting')
    initWebSocket()
  }
}, 3000)

/* WebSocket connection with exponential backoff */
function initWebSocket() {
  connect()
}

function connect() {
  var protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:'
  var host = window.location.host
  try {
    ws = new WebSocket(`${protocol}//${host}`)
  } catch (err) {
    console.error('WebSocket constructor error:', err)
    scheduleReconnect()
    return
  }

  ws.onopen = () => {
    reconnectAttempt = 0
    updateStatus('connected')
    updateBridgeUrl()
    console.log('WebSocket connected')
    var documentUrl = null
    try {
      documentUrl = Office.context.document.url || null
    } catch (_e) {
      // Not available in standalone/browser mode
    }
    ws.send(JSON.stringify({ type: 'ready', documentUrl: documentUrl }))

    // Enable auto-start on document open (shared runtime)
    try {
      if (Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) {
        Office.addin.setStartupBehavior(Office.StartupBehavior.load)
      }
    } catch (_e) {
      // SharedRuntime not available — manual button still works
    }

    // Tag document for auto-show taskpane
    try {
      Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true)
      Office.context.document.settings.saveAsync()
    } catch (_e) {
      // Settings API not available in standalone mode
    }
  }

  ws.onclose = () => {
    updateStatus('disconnected')
    console.log('WebSocket closed')
    scheduleReconnect()
  }

  ws.onerror = (err) => {
    console.error('WebSocket error:', err)
  }

  ws.onmessage = (event) => {
    try {
      const message = JSON.parse(event.data)
      handleCommand(message)
    } catch (err) {
      console.error('Failed to parse message:', err)
    }
  }
}

function scheduleReconnect() {
  var delay = Math.min(BASE_DELAY * 2 ** reconnectAttempt, MAX_DELAY)
  var jitter = Math.floor(Math.random() * 1000)
  reconnectAttempt++
  console.log(`Reconnecting in ${delay + jitter}ms (attempt ${reconnectAttempt})`)
  setTimeout(connect, delay + jitter)
}

/* Bridge URL display */
function updateBridgeUrl() {
  var el = document.getElementById('bridge-url')
  if (el) el.textContent = window.location.origin
}

/* Status display */
function updateStatus(state) {
  var el = document.getElementById('status')
  if (!el) return
  var labels = { connected: 'Connected', disconnected: 'Disconnected', connecting: 'Connecting...' }
  el.textContent = labels[state] || state
  el.className = `status ${state}`
}

/* Command queue — Office.js PowerPoint.run() calls must execute sequentially.
   Concurrent calls compete for the same COM context and can deadlock. */
var commandQueue = []
var commandRunning = false

function processQueue() {
  if (commandRunning || commandQueue.length === 0) return
  commandRunning = true
  var next = commandQueue.shift()
  next().finally(() => {
    commandRunning = false
    processQueue()
  })
}

/* Command handler — dispatches actions to execution functions */
function handleCommand(message) {
  if (message.type !== 'command' || !message.id) return

  if (message.action === 'executeCode') {
    commandQueue.push(() => executeCode(message.params.code, message.id))
    processQueue()
  } else {
    sendError(message.id, { message: `Unknown action: ${message.action}` })
  }
}

/* Execute a code string inside PowerPoint.run() via AsyncFunction */
function executeCode(code, requestId) {
  if (typeof PowerPoint === 'undefined') {
    sendError(requestId, {
      message: 'PowerPoint not available — running in standalone mode',
      code: 'NotInOffice',
    })
    return Promise.resolve()
  }

  return PowerPoint.run((context) => {
    var fn = new AsyncFunction('context', 'PowerPoint', code)
    return fn(context, PowerPoint)
  })
    .then((result) => {
      sendResponse(requestId, result === undefined ? null : result)
    })
    .catch((error) => {
      var errorObj = {
        message: error.message || String(error),
        code: error.code || 'UnknownError',
      }
      if (error.debugInfo) {
        errorObj.debugInfo = error.debugInfo
      }
      sendError(requestId, errorObj)
    })
}

/* Send a success response to the server */
function sendResponse(id, data) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify({ type: 'response', id: id, data: data }))
  }
}

/* Send an error response to the server */
function sendError(id, error) {
  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify({ type: 'error', id: id, error: error }))
  }
}
