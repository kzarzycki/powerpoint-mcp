import { describe, expect, it } from 'vitest'
import { withTool } from './tool-helpers.ts'

describe('withTool', () => {
  it('converts a thrown error into an isError text result', async () => {
    const wrapped = withTool(async () => {
      throw new Error('boom')
    })
    const result = await wrapped({}, {} as never)
    expect(result.isError).toBe(true)
    expect(result.content).toEqual([{ type: 'text', text: 'Error: boom' }])
  })

  it('stringifies a non-Error throw', async () => {
    const wrapped = withTool(async () => {
      throw 'plain string failure'
    })
    const result = await wrapped({}, {} as never)
    expect(result.isError).toBe(true)
    expect(result.content).toEqual([{ type: 'text', text: 'Error: plain string failure' }])
  })

  it('pretty-prints an object return as JSON text', async () => {
    const wrapped = withTool(async () => ({ a: 1, b: 'two' }))
    const result = await wrapped({}, {} as never)
    expect(result.isError).toBeUndefined()
    expect(result.content).toEqual([{ type: 'text', text: JSON.stringify({ a: 1, b: 'two' }, null, 2) }])
  })

  it('passes a string return through unchanged', async () => {
    const wrapped = withTool(async () => 'already formatted\n\nwith suffix')
    const result = await wrapped({}, {} as never)
    expect(result.isError).toBeUndefined()
    expect(result.content).toEqual([{ type: 'text', text: 'already formatted\n\nwith suffix' }])
  })

  it('passes a pre-built CallToolResult envelope through unchanged', async () => {
    const envelope = {
      content: [
        { type: 'image' as const, data: 'abc', mimeType: 'image/png' },
        { type: 'text' as const, text: 'caption' },
      ],
    }
    const wrapped = withTool(async () => envelope)
    const result = await wrapped({}, {} as never)
    expect(result).toEqual(envelope)
  })
})
