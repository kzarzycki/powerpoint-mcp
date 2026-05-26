# End-to-end tests

These tests exercise the full stack — MCP client → bridge server → Office.js add-in
running in **PowerPoint on the web** — by driving a real browser with Playwright.
Unlike the unit tests (`npm run check`), they need a Microsoft 365 account and a test
deck stored in OneDrive/SharePoint.

The browser uses a **persistent profile** (`e2e/.browser-profile`, gitignored) so you
sign in to M365 once. After that the suite runs headless and unattended.

## Prerequisites (one-time)

1. **Local TLS certs** — the bridge runs over HTTPS/WSS so WAC will load the add-in:
   ```bash
   npm run setup-certs        # needs mkcert: brew install mkcert && mkcert -install
   ```

2. **Playwright Chromium**:
   ```bash
   npx playwright install chromium
   ```

3. **A test deck in OneDrive/SharePoint.** Upload a `.pptx`, open it in PowerPoint on
   the web, and copy the edit URL from the address bar.

4. **`e2e/local-config.json`** (gitignored) pointing at that deck:
   ```json
   { "E2E_DOC_URL": "https://<tenant>-my.sharepoint.com/.../Doc.aspx?...&action=edit" }
   ```

## Step by step

### 1. Bootstrap the browser profile

```bash
npm run e2e:setup-profile
```

This **deletes any existing profile** and starts fresh. A Chromium window opens with
two tabs:

- **Tab 1** — a splash page with these instructions (flip back to it anytime).
- **Tab 2** — the test deck, which redirects to Microsoft 365 sign-in.

**Sign in to M365 in tab 2.** You have 10 minutes. Once the deck renders, the script
takes over automatically: it accepts the two sideload dialogs ("Enable Developer Mode",
"Registering Developer Add-in Manifest"), opens the add-in taskpane, and waits for the
bridge connection. On success the splash turns green ✓ and the window closes. The
profile is now primed.

You only repeat this when M365 cookies expire or you want a clean profile.

### 2. Run the tests

```bash
npm run test:e2e
```

Runs headless against the primed profile. The suite starts its own bridge + MCP server
(ports 9443 / 9001), navigates to the deck, waits for the add-in to connect, and drives
MCP tools end to end.

To watch a run in a visible browser:

```bash
E2E_HEADED=1 npm run test:e2e
```

To run a single test, pass a `-g` title filter through to Playwright:

```bash
npm run test:e2e -- -g "inspect_deck reports"
```

## How it works

- **`global-setup.ts`** starts the bridge server (`BRIDGE_TLS=1`, ports 9443/9001) and
  waits for both `/health` endpoints before any test runs.
- **`fixtures/pptx-page.ts`** opens a page per test on the persistent profile, with a
  spoofed Chrome UA (WAC skips sideloading for `HeadlessChrome`) and a CDP route shim
  that lets the public WAC origin reach the loopback bridge (Private Network Access).
- **`helpers/sideload.ts`** holds the shared connect logic (`connectAddin`): dismiss the
  sideload dialogs, give `AutoShowTaskpaneWithDocument` a short grace, then click the
  ribbon button if the pane didn't auto-open. The first navigation after a fresh manifest
  registration may not have the ribbon commands installed yet, so it reloads once and
  retries.
- Readiness is gated on the **bridge `/health` connection count**, not a DOM probe — the
  add-in taskpane is a nested iframe inside the WAC host frame and isn't reachable from a
  top-level locator.

## Troubleshooting

- **"TLS certs not found"** — run `npm run setup-certs`.
- **"E2E_DOC_URL not set"** — create `e2e/local-config.json` (above).
- **"port 9443/9001 already in use"** — a dev bridge is running; stop it first.
- **Tests redirect to login / time out connecting** — the profile's M365 cookies expired.
  Re-run `npm run e2e:setup-profile`.
- **Setup splash turned red ✗** — sign-in wasn't completed in time. Re-run and sign in
  promptly in tab 2.
