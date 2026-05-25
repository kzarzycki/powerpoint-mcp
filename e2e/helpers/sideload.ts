import type { BrowserContext } from '@playwright/test'
import { E2E_BRIDGE_PORT } from '../config.ts'
import { waitForAddinConnection } from './wait-for-connection.ts'

export type PptxPage = Awaited<ReturnType<BrowserContext['newPage']>>
export type AddinFrame = ReturnType<PptxPage['frames']>[number]

/**
 * Find the add-in taskpane frame. It loads from the bridge (localhost:PORT) but
 * is nested inside the WAC host frame, so a top-level locator/frameLocator can't
 * reach it — enumerate all frames instead.
 */
export function findAddinFrame(page: PptxPage): AddinFrame | undefined {
  return page.frames().find((f) => f.url().includes(`localhost:${E2E_BRIDGE_PORT}/index.html`))
}

/**
 * Poll every frame of the page and dismiss the two sideload gating dialogs Office
 * Web raises before an add-in loads: "Enable Developer Mode" (check the box, then
 * OK) and "Registering Developer Add-in Manifest" (Yes). Returns once the taskpane
 * frame appears, or shortly after the manifest dialog is accepted, or the window
 * elapses (long enough to also cover an interactive M365 sign-in beforehand).
 */
export async function dismissSideloadDialogs(page: PptxPage, windowMs = 60_000): Promise<void> {
  const deadline = Date.now() + windowMs
  let devModeHandled = false
  let manifestHandled = false
  // Once the manifest dialog is accepted, keep polling only briefly — the pane
  // either auto-opens (addin frame appears) or it won't, and we fall back to the
  // ribbon. Avoids burning the whole window when auto-show doesn't fire.
  let settleDeadline = Number.POSITIVE_INFINITY
  while (Date.now() < deadline && Date.now() < settleDeadline) {
    // Taskpane already loaded — nothing left to dismiss.
    if (findAddinFrame(page)) return

    for (const frame of page.frames()) {
      try {
        if (
          !devModeHandled &&
          (await frame
            .getByText(/enable developer mode/i)
            .first()
            .isVisible())
        ) {
          await frame.locator('input[type="checkbox"]').first().check()
          await frame.getByRole('button', { name: /^ok$/i }).first().click()
          devModeHandled = true
          console.log('[e2e] Enabled developer mode')
          continue
        }
        if (
          !manifestHandled &&
          (await frame
            .getByText(/registering developer add-?in manifest/i)
            .first()
            .isVisible())
        ) {
          await frame.getByRole('button', { name: /^yes$/i }).first().click()
          manifestHandled = true
          settleDeadline = Date.now() + 10_000
          console.log('[e2e] Accepted add-in manifest registration')
        }
      } catch {
        // Frame detached or dialog vanished mid-check — keep polling.
      }
    }
    await page.waitForTimeout(1000)
  }
}

/**
 * Open the add-in taskpane from the ribbon. After a fresh manifest registration,
 * Office Web does NOT honor AutoShowTaskpaneWithDocument — the pane only auto-opens
 * once it has been opened manually at least once on the document. The button lives
 * in the WAC host frame (euc-powerpoint.officeapps.live.com), labelled by the
 * manifest DisplayName. Returns true if a button was found and clicked.
 */
export async function openTaskpaneFromRibbon(page: PptxPage, windowMs = 30_000): Promise<boolean> {
  const deadline = Date.now() + windowMs
  while (Date.now() < deadline) {
    for (const frame of page.frames()) {
      try {
        const button = frame.getByRole('button', { name: /^powerpoint mcp$/i }).first()
        if (await button.isVisible()) {
          await button.click()
          console.log('[e2e] Opened taskpane from ribbon')
          return true
        }
      } catch {
        // Frame detached or button not in this frame — keep scanning.
      }
    }
    await page.waitForTimeout(1000)
  }
  return false
}

async function isConnected(timeoutMs: number): Promise<boolean> {
  try {
    await waitForAddinConnection(timeoutMs)
    return true
  } catch {
    return false
  }
}

/**
 * Try to connect on the currently-loaded page: give AutoShowTaskpaneWithDocument a
 * short grace (a primed document opens the pane and connects within a few seconds),
 * then fall back to clicking the ribbon button. The grace guard means an already-open
 * pane is never toggled shut.
 */
async function attemptConnect(page: PptxPage, dialogWindowMs: number): Promise<boolean> {
  await dismissSideloadDialogs(page, dialogWindowMs)
  if (await isConnected(8_000)) return true
  console.log('[e2e] Taskpane did not auto-open; opening from ribbon')
  if (await openTaskpaneFromRibbon(page)) {
    return isConnected(60_000)
  }
  return false
}

/**
 * Drive a freshly-navigated PowerPoint Web page to a connected add-in.
 *
 * The hard case is the very first navigation after a clean manifest registration:
 * Office Web installs the add-in's ribbon commands asynchronously, so neither
 * AutoShowTaskpaneWithDocument nor the ribbon button is available yet. A reload
 * picks up the now-cached commands — after which the pane auto-opens (or the ribbon
 * button is immediately clickable). So: try once, and if that fails, reload and try
 * again.
 *
 * @param dialogWindowMs how long to wait for the sideload dialogs to appear — bump
 *   this when an interactive M365 sign-in may precede them.
 */
export async function connectAddin(page: PptxPage, dialogWindowMs = 60_000): Promise<void> {
  if (await attemptConnect(page, dialogWindowMs)) return

  console.log('[e2e] Reloading to pick up freshly-registered add-in commands')
  await page.reload({ waitUntil: 'domcontentloaded', timeout: 60_000 })
  if (await attemptConnect(page, 20_000)) return

  throw new Error('Add-in did not connect: taskpane never auto-opened and the ribbon button never appeared')
}
