import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, realpathSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  claim,
  type LoopState,
  markImplemented,
  markMerged,
  recordGate,
  recordLive,
  review,
  setArtifact,
  waiveLive,
} from './engineering-loop.ts'
import { parseOmpOutput, parseReviewOutput } from './engineering-review.ts'
import { createState, readState, updateState } from './loop-store.ts'

function cleanEnvironment(): NodeJS.ProcessEnv {
  const inherited = Object.entries(process.env).filter(([key]) => !key.startsWith('GIT_'))
  // Temporary repositories must not depend on a host or CI identity.
  return {
    ...Object.fromEntries(inherited),
    GIT_AUTHOR_NAME: 'test',
    GIT_AUTHOR_EMAIL: 'test@example.com',
    GIT_COMMITTER_NAME: 'test',
    GIT_COMMITTER_EMAIL: 'test@example.com',
  }
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
  for (const clone of [first, second]) {
    git(clone, ['config', 'user.name', 'test'])
    git(clone, ['config', 'user.email', 'test@example.com'])
  }
  return { first, second }
}

function approveFile(dir: string, name = 'approve.json'): string {
  const path = join(dir, name)
  writeFileSync(path, JSON.stringify({ verdict: 'APPROVE', findings: [], reviewer: 'test' }))
  return path
}

function claimToImplemented(cwd: string, issue: number): LoopState {
  const approve = approveFile(cwd)
  let state = claim(cwd, String(issue), `test-${issue}`, false)
  writeFileSync(join(state.worktree, 'spec.md'), 'spec')
  setArtifact(cwd, issue, 'spec', join(state.worktree, 'spec.md'), false)
  review(cwd, issue, 'spec', undefined, false, approve)
  writeFileSync(join(state.worktree, 'plan.md'), 'plan')
  setArtifact(cwd, issue, 'plan', join(state.worktree, 'plan.md'), false)
  review(cwd, issue, 'plan', undefined, false, approve)
  git(state.worktree, ['commit', '--allow-empty', '-m', 'impl'])
  state = markImplemented(cwd, issue, false)
  return state
}

function claimToGatesGreen(cwd: string, issue: number): LoopState {
  claimToImplemented(cwd, issue)
  review(cwd, issue, 'branch', undefined, false, approveFile(cwd))
  const previousRunner = process.env.LOOP_GATE_RUNNER
  process.env.LOOP_GATE_RUNNER = ''
  try {
    return recordGate(cwd, issue, 'true', false)
  } finally {
    if (previousRunner === undefined) delete process.env.LOOP_GATE_RUNNER
    else process.env.LOOP_GATE_RUNNER = previousRunner
  }
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

  it('rejects a stale compare-and-swap update after a concurrent write', () => {
    const { first, second } = repository()
    const initial = { issue: 126, owner: 'test:first', step: 0 }
    const created = createState(first, 126, initial)
    const readByB = readState<typeof initial>(second, 126)
    if (!readByB) throw new Error('expected state to exist')
    updateState(second, 126, readByB.oid, { ...readByB.value, step: 1 })
    expect(() => updateState(first, 126, created.oid, { ...initial, step: 2 })).toThrow(
      /state conflict|state update lost/,
    )
    expect(readState(first, 126)?.value).toEqual({ ...initial, step: 1 })
  })

  it('creates the story worktree as a sibling of the repo, not nested inside it', () => {
    const { first } = repository()
    const previous = process.env.AGENT_SESSION
    process.env.AGENT_SESSION = 'test:sibling'
    try {
      const state = claim(first, '9001', 'sibling-check', false)
      expect(state.worktree).toBe(join(dirname(realpathSync(first)), 'powerpoint-mcp--issue-9001'))
      expect(existsSync(state.worktree)).toBe(true)
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
    }
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
      '{"type":"turn_end","message":{"role":"assistant","content":[{"type":"text","text":""}]}}',
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
      attempts: { spec: 0, plan: 0, branch: 0, checks: 0 },
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

  it('rejects a branch review when the worktree HEAD has moved past the implementation head', () => {
    const { first } = repository()
    const previous = process.env.AGENT_SESSION
    process.env.AGENT_SESSION = 'test:head-branch'
    try {
      const issue = 201
      const state = claimToImplemented(first, issue)
      git(state.worktree, ['commit', '--allow-empty', '-m', 'drift'])
      expect(() => review(first, issue, 'branch', undefined, false, approveFile(first))).toThrow(
        /has moved past the recorded implementation head/,
      )
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
    }
  })

  it('does not consume a checks attempt on a spawn-level infra failure, but does on a genuine failure', () => {
    const { first } = repository()
    const previous = process.env.AGENT_SESSION
    const previousRunner = process.env.LOOP_GATE_RUNNER
    process.env.AGENT_SESSION = 'test:head-checks'
    try {
      const issue = 202
      claimToImplemented(first, issue)
      review(first, issue, 'branch', undefined, false, approveFile(first))

      process.env.LOOP_GATE_RUNNER = '/nonexistent/loop-gate-runner-xyz'
      expect(() => recordGate(first, issue, 'true', false)).toThrow(/checks command failed to run/)
      const afterInfra = readState<LoopState>(first, issue)?.value
      expect(afterInfra?.attempts.checks).toBe(0)
      expect(afterInfra?.phase).toBe('BRANCH_APPROVED')

      process.env.LOOP_GATE_RUNNER = ''
      const afterFail = recordGate(first, issue, 'exit 1', false)
      expect(afterFail.attempts.checks).toBe(1)
      expect(afterFail.phase).toBe('BRANCH_APPROVED')

      const afterPass = recordGate(first, issue, 'true', false)
      expect(afterPass.attempts.checks).toBe(2)
      expect(afterPass.phase).toBe('GATES_GREEN')
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
      if (previousRunner === undefined) delete process.env.LOOP_GATE_RUNNER
      else process.env.LOOP_GATE_RUNNER = previousRunner
    }
  })

  it('reopens a branch-approved story to IMPLEMENTED only when HEAD actually moved', () => {
    const { first } = repository()
    const previous = process.env.AGENT_SESSION
    process.env.AGENT_SESSION = 'test:reopen'
    try {
      const issue = 203
      claimToImplemented(first, issue)
      const worktree = readState<LoopState>(first, issue)?.value.worktree as string
      const approved = review(first, issue, 'branch', undefined, false, approveFile(first))
      expect(approved.phase).toBe('BRANCH_APPROVED')

      const idempotent = markImplemented(first, issue, false)
      expect(idempotent.phase).toBe('BRANCH_APPROVED')

      git(worktree, ['commit', '--allow-empty', '-m', 'fix after branch review'])
      const reopened = markImplemented(first, issue, false)
      expect(reopened.phase).toBe('IMPLEMENTED')
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
    }
  })

  it('reopens a gates-green story to IMPLEMENTED and invalidates stale live evidence when HEAD moved', () => {
    const { first } = repository()
    const previous = process.env.AGENT_SESSION
    process.env.AGENT_SESSION = 'test:reopen-green'
    try {
      const issue = 206
      const gatesGreen = claimToGatesGreen(first, issue)
      const evidence = join(gatesGreen.worktree, 'evidence.txt')
      writeFileSync(evidence, 'live evidence')
      recordLive(first, issue, evidence, false)

      const idempotent = markImplemented(first, issue, false)
      expect(idempotent.phase).toBe('GATES_GREEN')
      expect(idempotent.live?.status).toBe('evidence')

      git(gatesGreen.worktree, ['commit', '--allow-empty', '-m', 'fix after checks'])
      const reopened = markImplemented(first, issue, false)
      expect(reopened.phase).toBe('IMPLEMENTED')
      expect(reopened.live).toBeUndefined()
      expect(() => markMerged(first, issue, 'https://example/pr/1', false)).toThrow(/is IMPLEMENTED/)
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
    }
  })

  it('requires live evidence or a waiver, and an unmoved head, before merging', () => {
    const { first } = repository()
    const previous = process.env.AGENT_SESSION
    process.env.AGENT_SESSION = 'test:merge'
    try {
      const issue = 204
      const gatesGreen = claimToGatesGreen(first, issue)
      expect(() => markMerged(first, issue, 'https://example/pr/1', false)).toThrow(
        /no recorded live evidence or waiver/,
      )

      git(gatesGreen.worktree, ['commit', '--allow-empty', '-m', 'drift after checks'])
      waiveLive(first, issue, 'no Office UI in this story', false)
      expect(() => markMerged(first, issue, 'https://example/pr/1', false)).toThrow(
        /has moved past the recorded implementation head/,
      )
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
    }
  })

  it('merges once live evidence and an unmoved head are both satisfied', () => {
    const { first } = repository()
    const previous = process.env.AGENT_SESSION
    process.env.AGENT_SESSION = 'test:merge-ok'
    try {
      const issue = 205
      const gatesGreen = claimToGatesGreen(first, issue)
      const evidence = join(gatesGreen.worktree, 'evidence.txt')
      writeFileSync(evidence, 'live evidence')
      recordLive(first, issue, evidence, false)
      const merged = markMerged(first, issue, 'https://example/pr/1', false)
      expect(merged.phase).toBe('MERGED')
      expect(merged.merge).toEqual({ pr: 'https://example/pr/1', commit: gatesGreen.implementation?.head })
    } finally {
      if (previous === undefined) delete process.env.AGENT_SESSION
      else process.env.AGENT_SESSION = previous
    }
  })
})
