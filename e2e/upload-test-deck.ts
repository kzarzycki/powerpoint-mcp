/**
 * Uploads the e2e test deck to OneDrive and captures the PowerPoint Web URL.
 *
 * Strategy: The M365 portal embeds OneDrive in an iframe. Instead of fighting that,
 * we find the SharePoint OneDrive URL and navigate there directly for the full UI.
 */

import { writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { chromium } from '@playwright/test'
import { BROWSER_PROFILE_DIR } from './config.ts'

const DECK_PATH = resolve(import.meta.dirname, 'fixtures/e2e-test-deck.pptx')

const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
  headless: false,
  ignoreHTTPSErrors: true,
  args: ['--disable-blink-features=AutomationControlled'],
})

const page = context.pages()[0] || (await context.newPage())

try {
  // Step 1: Discover the SharePoint OneDrive URL
  console.log('[1/5] Discovering SharePoint OneDrive URL...')
  await page.goto('https://m365.cloud.microsoft/onedrive/myfiles', {
    waitUntil: 'domcontentloaded',
    timeout: 30_000,
  })
  await page.waitForTimeout(6000)

  // Look for iframes — OneDrive is likely embedded in one
  const iframes = page.frames()
  console.log(`Found ${iframes.length} frames`)
  let sharePointUrl = ''
  for (const frame of iframes) {
    const url = frame.url()
    console.log(`  Frame: ${url.substring(0, 120)}`)
    if (url.includes('sharepoint.com') || url.includes('my.sharepoint')) {
      sharePointUrl = url
    }
  }

  // Also check for links to SharePoint in the page
  if (!sharePointUrl) {
    const links = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('a[href*="sharepoint"], iframe[src*="sharepoint"]')).map(
        (el) => (el as HTMLAnchorElement).href || (el as HTMLIFrameElement).src,
      )
    })
    console.log('SharePoint links found:', links)
    if (links.length > 0) sharePointUrl = links[0]
  }

  if (!sharePointUrl) {
    // Last resort: check the OneDrive API response for the user's OneDrive URL
    const apiUrl = await page.evaluate(async () => {
      try {
        const r = await fetch('https://graph.microsoft.com/v1.0/me/drive', {
          headers: { Authorization: 'Bearer ' },
        })
        const d = await r.json()
        return d.webUrl || ''
      } catch {
        return ''
      }
    })
    if (apiUrl) sharePointUrl = apiUrl
  }

  if (sharePointUrl) {
    // Extract the base SharePoint URL and navigate to it
    const spMatch = sharePointUrl.match(/(https:\/\/[^/]*sharepoint\.com\/[^?]*)/)
    if (spMatch) {
      const directUrl = spMatch[1]
      console.log(`[2/5] Navigating to SharePoint OneDrive: ${directUrl}`)
      await page.goto(directUrl, { waitUntil: 'domcontentloaded', timeout: 30_000 })
      await page.waitForTimeout(6000)
    }
  } else {
    console.log('Could not find SharePoint URL. Trying to work with embedded OneDrive...')
  }

  await page.screenshot({ path: '/tmp/e2e-step2-onedrive.png' })
  console.log('URL:', page.url())

  // List all buttons on the page for debugging
  const buttons = await page.getByRole('button').all()
  const btnInfo: string[] = []
  for (const btn of buttons.slice(0, 30)) {
    const label = await btn.getAttribute('aria-label').catch(() => null)
    const text = (await btn.textContent().catch(() => ''))?.trim()
    btnInfo.push(label || text || '(unnamed)')
  }
  console.log('Buttons (first 30):', btnInfo.join(' | '))

  // Step 3: Upload
  console.log('[3/5] Uploading...')
  let uploaded = false

  // Try "Upload" button directly
  for (const selector of [
    page.getByRole('button', { name: /^upload$/i }),
    page.getByRole('menuitem', { name: /upload/i }),
    page.locator('button:has-text("Upload")'),
    page.locator('[data-automationid="uploadCommand"]'),
    page.locator('[aria-label="Upload"]'),
    // For the new OneDrive: try "Add new" then "Files upload"
    page.getByRole('button', { name: /add new/i }),
    page.locator('[data-automationid="newCommand"]'),
  ]) {
    try {
      if (await selector.first().isVisible({ timeout: 2000 })) {
        await selector.first().click()
        console.log('Clicked upload trigger')
        await page.waitForTimeout(1500)

        // Check if a dropdown appeared with "Files upload"
        const fileUploadItem = page.getByRole('menuitem', { name: /file/i }).first()
        if (await fileUploadItem.isVisible({ timeout: 2000 }).catch(() => false)) {
          const [fc] = await Promise.all([
            page.waitForEvent('filechooser', { timeout: 10_000 }),
            fileUploadItem.click(),
          ])
          await fc.setFiles(DECK_PATH)
          uploaded = true
          break
        }

        // Maybe the button itself triggers file chooser directly
        // (in some OneDrive versions, "Upload" directly opens file dialog)
        break
      }
    } catch {}
  }

  if (!uploaded) {
    // Try keyboard shortcut or drag-and-drop approach
    // Fall back to finding file input element
    const fileInput = page.locator('input[type="file"]')
    if ((await fileInput.count()) > 0) {
      await fileInput.first().setInputFiles(DECK_PATH)
      uploaded = true
      console.log('Uploaded via hidden file input')
    }
  }

  if (!uploaded) {
    await page.screenshot({ path: '/tmp/e2e-upload-failed.png' })
    throw new Error('Could not find upload mechanism. Check /tmp/e2e-upload-failed.png and buttons list above.')
  }

  console.log('File upload initiated.')
  await page.waitForTimeout(8000)
  await page.screenshot({ path: '/tmp/e2e-step3-uploaded.png' })

  // Step 4: Open in PowerPoint Web
  console.log('[4/5] Opening in PowerPoint Web...')
  const fileItem = page.getByText('e2e-test-deck').first()
  await fileItem.waitFor({ state: 'visible', timeout: 15_000 })
  await fileItem.dblclick()
  await page.waitForTimeout(10_000)

  const finalPages = context.pages()
  const pptxPage = finalPages[finalPages.length - 1]
  const docUrl = pptxPage.url()

  console.log('[5/5] Capturing URL...')
  console.log('')
  console.log('=== PowerPoint Web URL ===')
  console.log(docUrl)
  console.log('')

  const envPath = resolve(import.meta.dirname, '.env')
  writeFileSync(envPath, `E2E_DOC_URL=${docUrl}\n`)
  console.log(`Written to ${envPath}`)
  await pptxPage.screenshot({ path: '/tmp/e2e-step4-pptx.png' })
} catch (err) {
  const msg = err instanceof Error ? err.message : String(err)
  console.error('Automation failed:', msg)
  console.log('Browser stays open for 60s. Close to exit early.')
  await Promise.race([new Promise((r) => setTimeout(r, 60_000)), new Promise((r) => context.on('close', r))])
}

await context.close()
console.log('Done.')
