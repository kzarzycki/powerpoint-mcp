import { describe, expect, it } from 'vitest'
import { toLocalFilePath } from './shared.ts'

describe('toLocalFilePath', () => {
  it('converts a file:// URL to a filesystem path', () => {
    expect(toLocalFilePath('file:///Users/me/Deck.pptx')).toBe('/Users/me/Deck.pptx')
  })

  it('decodes percent-encoded characters like spaces', () => {
    expect(toLocalFilePath('file:///Users/me/My%20Deck.pptx')).toBe('/Users/me/My Deck.pptx')
  })

  it('leaves a plain filesystem path untouched', () => {
    expect(toLocalFilePath('/Users/me/Deck.pptx')).toBe('/Users/me/Deck.pptx')
  })
})
