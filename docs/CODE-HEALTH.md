# Code health — findings and remediation decisions

Backlog from adversarial code review and session feedback. Evidence, severity, and release disposition are separate: a confirmed capability can be accepted under the local-trust policy; a proposed remedy can remain unproven. CH05 and CH10 retain their **PLAUSIBLE** flags. Feeds the [E00 epic](https://github.com/kzarzycki/powerpoint-mcp/issues/125) and the reliable-bridge release.

**24 findings** (16 originally rated major, 8 minor), not 24 reproduced defects. CH01–CH22 originate in the code review; CH23–CH24 in [session feedback](feedback/session-analysis-2026-07.md). Explicit verdicts below govern the challenged findings. Other entries retain their prior evidence and need a failing reproduction before implementation. Source line numbers are approximate anchors.

## Bridge & protocol

### CH01 · MAJOR — Command timeout starts at send-time while the add-in runs commands serially → false timeouts, silently dropped late responses, and executed-but-reported-failed mutations
`server/bridge.ts:121`

**Verdict:** reproduced with the actual bridge and add-in queue under Node 24.18.0 and a simulated Office host: two commands taking 250 ms each, a 400 ms timeout, first call fulfilled, second timed out, both mutations applied. This proves queue-wait ambiguity, not real-host timing. Relaunch blocker.

**Failure:** sendCommand starts the 30s timer the instant it calls targetWs.send() (bridge.ts:120-127). But the add-in executes every command strictly one-at-a-time behind a commandRunning flag / commandQueue (app.js:122-130): a new executeCode only starts after the previous PowerPoint.run() fully resolves. So the timer measures queue-wait + execution, not execution. Two ways this fires: (a) a single heavy command on a large deck (e.g. verify_slides or a big execute_officejs) that legitimately takes >30s — no concurrency needed; (b) concurrent MCP calls (multiple sessions — the code explicitly supports this via getConcurrentWarning — or parallel tool calls). N commands sent near-simultaneously are queued; command #k waits (k-1)*execTime before even starting, so later ones blow the 30s budget. On timeout, pendingRequests.delete(id) runs and the promise rejects with 'Command timed out' (bridge.ts:122-123). The add-in, however, keeps going and DOES execute the command (mutating the live deck), then sends its response — which handleResponse drops silently because the id is gone (bridge.ts:72-73 `if (!pending) return`). Net effect: the MCP tool reports failure/timeout for an operation that actually succeeded. An agent that retries (add_slide, insert shape, etc.) produces DUPLICATE slides/shapes. Multi-round-trip tools (copy_slides export then insert, shared.ts getLocalCopyPath revision-check at default 30s followed by 120s export) are especially exposed because the cheap first round-trip inherits the 30s default yet can sit behind a long queue.

**Remedy contract:** prefer server-side per-connection scheduling if it preserves compatibility with the shipped add-in. Execution timeout must not include time spent waiting behind another command. Neither a timeout nor a socket reconnect proves cancellation: the add-in queue belongs to the page and may continue after reconnect. Define bounded waiting and recovery when completion never arrives; do not release the execution slot merely because the caller timed out, or hold queued callers forever. Distinguish known-not-started from unknown-outcome failures, retain late-completion evidence without logging presentation content, and never automatically replay an ambiguous mutation. A `started` acknowledgement is an alternative requiring explicit protocol/version compatibility. Acceptance belongs to E00-02; changing error text alone does not close this finding.

### CH02 · MAJOR — Unsaved presentation identity changes on registration; duplicate registrations can leave stale pool entries
`server/bridge.ts:49-57,84-86`; `server/index.ts:396-415`

**Verdict:** identity instability is confirmed in source. Registering the same socket twice with `generateId(null)` reproduces two entries; `removeBySocket` removes only one. That boundary case is not the normal shipped-client path, which sends one `ready` per socket. CH03 is a reachable source of two sockets from one runtime. Sleep-related stale connections remain host-dependent, not reproduced.

**Failure:** every null document URL receives a new `untitled-N` key. Cached identifiers therefore fail after re-registration. Multiple entries block calls that omit `presentationId`; duplicate registrations for one socket can survive its close handler. Saved decks use the URL as a key, but saving or recreating a taskpane introduces different identity transitions.

**Remedy contract:** define identities across same-runtime reconnect, taskpane recreation, Save/Save As, and genuinely distinct open decks before choosing storage. A per-runtime UUID only solves the first boundary. Persisting identity in document settings requires real-host evidence about availability, copied decks, and document dirty state. Do not blindly close the replaced socket: its automatic reconnect can create replacement churn, and its mutations may still run. Repeated registration and cleanup must be idempotent without dropping another live deck. Specify this separately from the CH03 pilot.

### CH03 · MAJOR — One add-in runtime can open two WebSocket connections
`addin/app.js:11-41,94-100`

**Verdict:** reproduced by executing the actual `addin/app.js` in an isolated VM: the standalone fallback followed by delayed Office readiness constructs two sockets. Pilot story for the engineering loop ([E00-01](https://github.com/kzarzycki/powerpoint-mcp/issues/127)).

**Failure:** `Office.onReady` and the 3-second fallback both call `initWebSocket()` / `connect()`. The latter unconditionally constructs a socket. Executing the actual `addin/app.js` in an isolated VM with fallback first and Office readiness second produces two socket constructions. Real cold-start frequency is unmeasured. With no document URL, the two registrations receive different pool keys; with a saved URL, one entry replaces the other, so the pool count does not necessarily reveal both sockets. This proves duplicate connections, not duplicate execution of every command.

**Remedy contract:** exactly one connection attempt may be in flight or open per runtime, regardless of readiness ordering, and no stale timer may add another. Reconnection after a closed socket must still work, with backoff preserved. Do not suppress reconnects by treating a closing or closed socket as live. Acceptance and live evidence are specified in [E00-01](https://github.com/kzarzycki/powerpoint-mcp/issues/127).

### CH04 · MINOR — Connection liveness after host suspension is unproven
`server/index.ts:378-420`

**Verdict:** no server heartbeat exists; the claimed stale-connection failure is host-dependent and unproven. Defer heartbeat implementation from the pilot, not investigation of recovery.

**Risk:** the bridge uses loopback, so Wi-Fi path loss is not evidence of a broken add-in socket. Loopback also does not prove every host/stack suspension mode impossible. A protocol pong may be serviced while Office execution is stuck; transport liveness is not Office readiness.

**Evidence required:** observe socket closure, pool membership, command outcomes, and recovery on Mac sleep/wake, taskpane recreation, and PowerPoint termination. Add heartbeat only if it detects an observed failure and passes false-disconnect checks after suspension. Do not use it as a substitute for CH01's command-outcome contract.

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

### CH16 · MAJOR — WebSocket registration lacks Origin validation
`server/index.ts:378-407`; `server/bridge.ts:40-42,71-85`

**Verdict:** missing validation is confirmed in source. A client admitted to the socket can register an arbitrary document URL, replace its pool entry, and submit response frames. Foreign-page reachability depends on browser/host network policy and needs an isolated browser reproduction.

**Risk:** loopback binding does not distinguish a trusted local client from a browser page able to reach the port. WebSocket access must not be inferred safe from HTTP CORS. Conversely, the static handler's OPTIONS headers do not prove every browser permits a WebSocket upgrade.

**Remedy contract:** observe actual Origin/Host headers from PowerPoint for Mac and PowerPoint Web, including TLS and custom ports. Reject foreign origins and malformed registration without breaking those clients. Decide absent/`null` Origin handling from that evidence, not a guessed sandbox exemption. Verify rejected connections cannot alter the pool, replace a target, or settle its requests. Authentication remains a separate local-trust policy decision; an Origin check is not authentication.

### CH17 · MAJOR — MCP HTTP grants live-deck control to local clients without credentials
`server/index.ts:230-271,444-490`

**Verdict:** confirmed capability, accepted for the documented trusted single-user, loopback-only deployment. Authentication is deferred for this release scope, not declared unnecessary for other deployments.

**Boundary:** Host validation and random session IDs are not credentials. A local client can initialize its own session. Browser control is a separate question: test Host, Origin, preflight, content-type, and session behavior rather than equating a local client with a foreign page. Shared-user and remote deployments remain unsupported.

**Disposition:** keep the security contract in the root README accurate. Revisit token provisioning, client compatibility, and secret/log handling before changing the trust model or exposing a non-loopback listener. Do not introduce authentication without a tested installation path for every supported client.

### CH18 · MINOR — Arbitrary Office.js execution has no enforced network-egress policy
`addin/index.html:3-9`; `addin/app.js:144-169`

**Verdict:** defense-in-depth gap within the trusted-client execution capability. Enforced CSP is deferred for the local-only relaunch; this does not make untrusted code safe.

**Risk:** `AsyncFunction` runs with host APIs, including available network APIs. A restrictive `default-src 'self'` policy would also block the CDN Office.js script and dynamic code execution unless explicitly allowed; a `connect-src` limited to the Office CDN would omit the bridge WebSocket. A CSP is not by itself a complete sandbox.

**Disposition:** preserve the documented trusted-client boundary. If egress restriction becomes a product requirement, evaluate a report-only policy in both Mac and Web hosts before enforcement, including dynamic execution, Office loading, bridge reconnect, and supported media operations.

### CH19 · MINOR — parseJsonBody buffers unbounded request bodies before any auth/session check
`server/index.ts:208`

**Failure:** handleMcpPost calls `parseJsonBody` (index.ts:208-221) which concatenates all `data` chunks into memory with no size cap, and this happens (index.ts:232) before the session/initialize gate at index.ts:235-258. A local client (or a browser hitting the correct Host) can POST a multi-gigabyte body to /mcp and the server buffers the whole thing into `chunks` before rejecting, exhausting memory / crashing the process — a pre-auth memory DoS. Lower severity because :3001 is loopback-bound, but it is reachable by any local process and by same-origin add-in pages.

**Remedy contract:** enforce a byte limit before concatenation and JSON parsing. Derive the limit from supported HTTP requests, including base64 images; do not choose an arbitrary 1–4 MB or 32 MiB cap. Oversized requests must stop accumulating memory, produce a documented client-visible failure, and leave the server responsive. Cover chunked bodies, disconnect during upload, and legitimate near-limit operations. Do not promise an HTTP 413 after destroying the response socket.

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

## Session-mined additions (2026-07)

### CH23 · MAJOR — edit_slide_zip returns success for edits the reimport path silently discards (presentation.xml and every other deck-level part)
`server/tools/xml.ts:141`

**Failure:** The zip round-trip is single-slide: export is `slide.exportAsBase64()` (a one-slide package whose `presentation.xml` is synthesized and whose slide part is renamed `slide1.xml`), reimport is `insertSlidesFromBase64`, which cannot carry deck-level state into the host deck. `edit_slide_zip` accepts arbitrary zip paths with no validation and unconditionally returns `{success: true}` (tools/xml.ts:141-144), so an edit to `presentation.xml` — e.g. `<p:sldId show="0">` to hide a slide — reports success while the change is thrown away on reimport. Observed cost in the field: a hide-slide task failed across multiple turns with no error anywhere, was root-caused by an agent spelunking the export format, and the session ended in user frustration (F04 in the session analysis; sections, theme application and version tracking hit the same wall).

**Fix:** Reject edits to paths the reimport cannot persist with a typed error that names the ceiling ("deck-level parts are regenerated on reimport; only slideN.xml content round-trips"), and document the export model in the tool description. Longer term E13-FB1 adds a real deck-level write path. Verify: **V2** (positive fixture: presentation.xml edit → typed error; negative: slide-part edit still round-trips), **V1**.

### CH24 · MAJOR — add_slide resolves layouts by name across all masters — multi-master decks get the wrong layout (issue #120)
`server/tools/slides.ts`

**Failure:** Layout resolution takes the first name match across every master, so in decks with multiple masters that reuse layout names (corporate templates — user: "not so rare in my company case") `add_slide` inserts the wrong master's layout: placeholder text drops into wrong shapes and shapes get wrong semantic names. Filed as GitHub issue #120 (open). Related latent defect: `extractThemeFromZip` (xml-helpers.ts:348-360) grabs the first `ppt/theme/*` part, which is not necessarily the applied theme — any brand/contrast logic built on it (E13/E14) inherits wrong colors/fonts; extends CH15.

**Fix:** Resolve layouts by (master, layout) id — or require/accept a master qualifier and error on ambiguous names; resolve the applied theme via the master's theme relationship instead of first-file. Verify: **V2** (multi-master fixture: ambiguous name → correct layout or typed ambiguity error; clean fixture unchanged), **V3**.
