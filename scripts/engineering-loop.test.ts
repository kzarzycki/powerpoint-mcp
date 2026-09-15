import { execFileSync } from 'node:child_process'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { type LoopState, review } from './engineering-loop.ts'
import { parseOmpOutput, parseReviewOutput } from './engineering-review.ts'
import { createState, readState } from './loop-store.ts'

function cleanEnvironment(): NodeJS.ProcessEnv {
  return Object.fromEntries(Object.entries(process.env).filter(([key]) => !key.startsWith('GIT_')))
}

function git(cwd: string, args: string[]): string {
  return execFileSync('git', args, { cwd, encoding: 'utf8', env: cleanEnvironment() }).trim()
}

function repository(): { first: string; second: string } {
  const base = mkdtempSync(join(tmpdir(), 'loop-base-'))
  const remote = join(base, 'remote.git')
  const seed = join(base, 'seed')
  git(base, ['init', seed])
  git(seed, ['config', 'user.name', 'test'])
  git(seed, ['config', 'user.email', 'test@example.com'])
  execFileSync('git', ['init', '--bare', remote], { cwd: base, encoding: 'utf8', env: cleanEnvironment() })
  writeFileSync(join(seed, 'README'), 'seed\n')
  git(seed, ['add', 'README'])
  git(seed, ['commit', '-m', 'seed'])
  git(seed, ['branch', '-M', 'main'])
  git(seed, ['remote', 'add', 'origin', remote])
  git(seed, ['push', 'origin', 'main'])
  const first = join(base, 'first')
  const second = join(base, 'second')
  git(base, ['clone', remote, first])
  git(base, ['clone', remote, second])
  return { first, second }
}

describe('engineering loop state', () => {
  it('allows exactly one competing remote claim', () => {
    const { first, second } = repository()
    const value = { issue: 126, owner: 'test:first' }
    expect(createState(first, 126, value).value).toEqual(value)
    expect(readState(second, 126)?.value).toEqual(value)
    expect(() => createState(second, 126, { issue: 126, owner: 'test:second' })).toThrow(
      /exclusive claim failed|exclusive claim lost/,
    )
    expect(readState(first, 126)?.value).toEqual(value)
  })

  it('keeps review parsing fail-closed', () => {
    expect(parseReviewOutput('{"verdict":"REVISE","findings":["missing test"]}')).toEqual({
      verdict: 'REVISE',
      findings: ['missing test'],
    })
    expect(() => parseReviewOutput('{"verdict":"APPROVE","findings":[],"reviewer":"author"}')).toThrow()
    expect(() => parseReviewOutput('{"verdict":"APPROVE","findings":[3]}')).toThrow()
  })

  it('extracts the verdict from the omp json event stream', () => {
    const stream = [
      '{"type":"message_update","assistantMessageEvent":{"type":"text_delta","delta":"{\\"verdict\\""}}',
      '{"type":"turn_end","message":{"role":"assistant","content":[{"type":"thinking","thinking":"x"},{"type":"text","text":"{\\"verdict\\":\\"APPROVE\\",\\"findings\\":[]}"}]}}',
      '{"type":"advisor_yielded"}',
    ].join('\n')
    expect(parseOmpOutput(stream)).toEqual({ verdict: 'APPROVE', findings: [] })
    expect(() => parseOmpOutput('{"type":"turn_start"}\n{"type":"advisor_yielded"}')).toThrow()
  })

  it('parks after three review rejections', () => {
    const { first } = repository()
    const state: LoopState = {
      issue: 126,
      issueUrl: 'https://github.com/kzarzycki/powerpoint-mcp/issues/126',
      owner: 'test:owner',
      claimNonce: 'test',
      branch: 'loop/test',
      worktree: first,
      phase: 'SPEC',
      attempts: { spec: 0, plan: 0, branch: 0, checks: 0, live: 0 },
      artifacts: {},
      reviews: [],
      gates: [],
      events: [],
      updatedAt: new Date().toISOString(),
    }
    createState(first, 126, state)
    const result = join(first, 'review.json')
    writeFileSync(result, '{"verdict":"REVISE","findings":["not ready"],"reviewer":"test"}')
    const previous = process.env.AGENT_SESSION
    process.env.AGENT_SESSION = 'test:owner'
    try {
      review(first, 126, 'spec', result, false, result)
      review(first, 126, 'spec', result, false, result)
      const parked = review(first, 126, 'spec', result, false, result)
      expect(parked.phase).toBe('PARKED')
      expect(parked.attempts.spec).toBe(3)
      expect(parked.reviews.filter((item) => item.verdict === 'REVISE')).toHaveLength(3)
      expect(parked.events.filter((item) => item.kind === 'REVIEW_REJECTED')).toHaveLength(3)
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
    }
  })
})
