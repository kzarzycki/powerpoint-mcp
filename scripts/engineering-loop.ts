#!/usr/bin/env node

import { execFileSync, spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, resolve } from 'node:path'
import { parseReviewOutput, type ReviewResult, runReview } from './engineering-review.ts'
import { createState, readState, type StateResult, updateState } from './loop-store.ts'

export const PHASES = [
  'TRIAGED',
  'SPEC',
  'SPEC_APPROVED',
  'PLAN',
  'PLAN_APPROVED',
  'IMPLEMENTED',
  'BRANCH_APPROVED',
  'GATES_GREEN',
  'MERGED',
  'PARKED',
] as const
export type Phase = (typeof PHASES)[number]
export type Gate = 'spec' | 'plan' | 'branch' | 'checks' | 'live'

export interface LoopEvent {
  at: string
  kind: string
  phase: Phase
  detail: string
  attempt?: number
}

export interface LoopState {
  issue: number
  issueUrl: string
  owner: string
  claimNonce: string
  branch: string
  worktree: string
  phase: Phase
  attempts: Record<Gate, number>
  artifacts: Partial<Record<'spec' | 'plan', { path: string; sha256: string }>>
  reviews: Array<{ gate: Gate; verdict: 'APPROVE' | 'REVISE'; reviewer: string; findings: string[]; at: string }>
  gates: Array<{ gate: Gate; command: string; exitCode: number; output: string; at: string }>
  events: LoopEvent[]
  updatedAt: string
}

const MAX_ATTEMPTS = 3
const EMPTY_ATTEMPTS: Record<Gate, number> = { spec: 0, plan: 0, branch: 0, checks: 0, live: 0 }

function now(): string {
  return new Date().toISOString()
}
function session(): string {
  const value = process.env.AGENT_SESSION
  if (!value) throw new Error('AGENT_SESSION is required for loop mutations')
  return value
}
function git(cwd: string, args: string[]): string {
  return execFileSync('git', args, { cwd, encoding: 'utf8' }).trim()
}
function root(cwd: string): string {
  return git(cwd, ['rev-parse', '--show-toplevel'])
}
function issueNumber(value: string): number {
  const match = value.match(/(?:issues\/|#)?(\d+)(?:$|[^\d])/)
  if (!match) throw new Error(`invalid issue: ${value}`)
  return Number(match[1])
}
function issueUrl(value: string, issue: number): string {
  return value.startsWith('http') ? value : `https://github.com/kzarzycki/powerpoint-mcp/issues/${issue}`
}
function slug(value: string): string {
  const result = value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '')
  if (!result) throw new Error('branch slug is empty')
  return result
}
function sha256(path: string): string {
  return createHash('sha256').update(readFileSync(path)).digest('hex')
}
function output(result: StateResult<LoopState>): LoopState {
  return result.value
}
function load(cwd: string, issue: number): StateResult<LoopState> {
  const state = readState<LoopState>(cwd, issue)
  if (!state) throw new Error(`issue #${issue} has no claim state`)
  return state
}
function update(cwd: string, state: StateResult<LoopState>, value: LoopState): StateResult<LoopState> {
  value.updatedAt = now()
  return updateState(cwd, value.issue, state.oid, value)
}
function event(value: LoopState, kind: string, detail: string, attempt?: number): LoopState {
  value.events.push({ at: now(), kind, phase: value.phase, detail, ...(attempt === undefined ? {} : { attempt }) })
  return value
}
function comment(issue: string, body: string): void {
  const result = spawnSync('gh', ['issue', 'comment', issue, '--body', body], { encoding: 'utf8' })
  if (result.status !== 0) throw new Error(`GitHub issue comment failed: ${result.stderr || result.stdout}`)
}
function record(
  cwd: string,
  state: StateResult<LoopState>,
  value: LoopState,
  publish: boolean,
): StateResult<LoopState> {
  const saved = update(cwd, state, value)
  if (publish) comment(value.issueUrl, `engineering-loop: ${value.phase}\n\n${value.events.at(-1)?.detail ?? ''}`)
  return saved
}
function requirePhase(value: LoopState, expected: Phase): void {
  if (value.phase !== expected) throw new Error(`issue #${value.issue} is ${value.phase}; expected ${expected}`)
}
function gate(value: LoopState, name: Gate): number {
  const next = value.attempts[name] + 1
  if (next > MAX_ATTEMPTS) throw new Error(`${name} gate is exhausted; issue is ${value.phase}`)
  value.attempts[name] = next
  return next
}
function parseArgs(args: string[]): { command: string; values: Record<string, string>; flags: Set<string> } {
  const [command = 'status', ...rest] = args
  const values: Record<string, string> = {}
  const flags = new Set<string>()
  for (let index = 0; index < rest.length; index += 1) {
    const token = rest[index]
    if (!token.startsWith('--')) continue
    const key = token.slice(2)
    const next = rest[index + 1]
    if (next && !next.startsWith('--')) {
      values[key] = next
      index += 1
    } else flags.add(key)
  }
  return { command, values, flags }
}
function required(values: Record<string, string>, key: string): string {
  const value = values[key]
  if (!value) throw new Error(`--${key} is required`)
  return value
}

export function claim(cwd: string, rawIssue: string, branchSlug: string, publish = true): LoopState {
  const issue = issueNumber(rawIssue)
  const url = issueUrl(rawIssue, issue)
  const repo = root(cwd)
  const branch = `loop/issue-${issue}-${slug(branchSlug)}`
  const worktree = resolve(dirname(repo), `${basename(repo)}--issue-${issue}`)
  const value: LoopState = {
    issue,
    issueUrl: url,
    owner: session(),
    claimNonce: `${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    branch,
    worktree,
    phase: 'TRIAGED',
    attempts: { ...EMPTY_ATTEMPTS },
    artifacts: {},
    reviews: [],
    gates: [],
    events: [],
    updatedAt: now(),
  }
  let created: StateResult<LoopState>
  try {
    created = createState(repo, issue, value)
  } catch (error) {
    const auditDir = resolve(repo, '.loop')
    mkdirSync(auditDir, { recursive: true })
    writeFileSync(
      resolve(auditDir, `rejected-claim-${value.claimNonce}.json`),
      JSON.stringify(
        { issue, owner: value.owner, branch, reason: error instanceof Error ? error.message : String(error) },
        null,
        2,
      ),
    )
    throw error
  }
  try {
    git(repo, ['worktree', 'add', '-b', branch, worktree, 'origin/main'])
  } catch (error) {
    value.phase = 'PARKED'
    event(value, 'PARKED', `worktree creation failed: ${error instanceof Error ? error.message : String(error)}`)
    update(repo, created, value)
    throw error
  }
  event(value, 'CLAIMED', `owner=${value.owner} branch=${branch} worktree=${worktree}`)
  const saved = update(repo, created, value)
  if (publish) comment(url, `engineering-loop: TRIAGED\n\nExclusive claim won by ${value.owner}; worktree ${worktree}.`)
  return output(saved)
}

export function setArtifact(
  cwd: string,
  issue: number,
  kind: 'spec' | 'plan',
  path: string,
  publish = true,
): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  if (kind === 'spec') {
    if (value.phase === 'TRIAGED') value.phase = 'SPEC'
    else requirePhase(value, 'SPEC')
  } else {
    if (value.phase === 'SPEC_APPROVED') value.phase = 'PLAN'
    else requirePhase(value, 'PLAN')
  }
  const artifactPath = resolve(cwd, path)
  if (!existsSync(artifactPath)) throw new Error(`artifact does not exist: ${artifactPath}`)
  value.artifacts[kind] = { path: artifactPath, sha256: sha256(artifactPath) }
  event(value, `${kind.toUpperCase()}_WRITTEN`, `${kind}=${artifactPath}`)
  return output(record(cwd, current, value, publish))
}

export function review(
  cwd: string,
  issue: number,
  kind: Gate,
  briefPath: string,
  publish = true,
  resultPath?: string,
): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  const expected: Record<Gate, Phase> = {
    spec: 'SPEC',
    plan: 'PLAN',
    branch: 'IMPLEMENTED',
    checks: 'BRANCH_APPROVED',
    live: 'GATES_GREEN',
  }
  requirePhase(value, expected[kind])
  const attempt = gate(value, kind)
  record(cwd, current, event(value, 'REVIEW_RESERVED', `${kind} attempt ${attempt}`, attempt), false)
  let result: ReviewResult
  try {
    if (resultPath) {
      const raw = JSON.parse(readFileSync(resolve(cwd, resultPath), 'utf8')) as Record<string, unknown>
      if (typeof raw.reviewer !== 'string') throw new Error('recorded review has no reviewer')
      result = {
        ...parseReviewOutput(JSON.stringify({ verdict: raw.verdict, findings: raw.findings })),
        reviewer: raw.reviewer,
      }
    } else {
      result = runReview({ cwd: value.worktree, brief: readFileSync(resolve(cwd, briefPath), 'utf8') })
    }
  } catch (error) {
    const failed = output(load(cwd, issue))
    event(failed, 'REVIEW_FAILED', `${kind}: ${error instanceof Error ? error.message : String(error)}`, attempt)
    record(cwd, load(cwd, issue), failed, publish)
    throw error
  }
  const latest = load(cwd, issue)
  const reviewed = output(latest)
  reviewed.reviews.push({
    gate: kind,
    verdict: result.verdict,
    reviewer: result.reviewer,
    findings: result.findings,
    at: now(),
  })
  if (result.verdict === 'APPROVE') {
    const next: Record<Gate, Phase> = {
      spec: 'SPEC_APPROVED',
      plan: 'PLAN_APPROVED',
      branch: 'BRANCH_APPROVED',
      checks: 'GATES_GREEN',
      live: 'GATES_GREEN',
    }
    reviewed.phase = next[kind]
    event(reviewed, 'REVIEW_APPROVED', `${kind} approved by ${result.reviewer}`, attempt)
  } else {
    event(reviewed, 'REVIEW_REJECTED', `${kind} rejected by ${result.reviewer}: ${result.findings.join('; ')}`, attempt)
    if (attempt >= MAX_ATTEMPTS) {
      reviewed.phase = 'PARKED'
      event(reviewed, 'PARKED', `${kind} rejected three times`, attempt)
    }
  }
  return output(record(cwd, latest, reviewed, publish))
}

export function recordGate(cwd: string, issue: number, command: string, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  requirePhase(value, 'BRANCH_APPROVED')
  const attempt = gate(value, 'checks')
  const result = spawnSync('mise', ['exec', 'node@24.18.0', '--', 'sh', '-lc', command], {
    cwd: value.worktree,
    encoding: 'utf8',
  })
  const text = `${result.stdout ?? ''}${result.stderr ?? ''}`.slice(-12000)
  value.gates.push({ gate: 'checks', command, exitCode: result.status ?? 1, output: text, at: now() })
  event(
    value,
    result.status === 0 ? 'GATE_PASSED' : 'GATE_FAILED',
    `checks attempt ${attempt}: exit ${result.status ?? 1}; ${command}`,
    attempt,
  )
  if (result.status === 0) value.phase = 'GATES_GREEN'
  else if (attempt >= MAX_ATTEMPTS) {
    value.phase = 'PARKED'
    event(value, 'PARKED', 'checks rejected three times', attempt)
  }
  return output(record(cwd, current, value, publish))
}

export function recordLive(cwd: string, issue: number, evidencePath: string, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  requirePhase(value, 'GATES_GREEN')
  const path = resolve(cwd, evidencePath)
  if (!existsSync(path)) throw new Error(`live evidence does not exist: ${path}`)
  event(value, 'LIVE_EVIDENCE', `operator evidence=${path}; live Office prerequisite must be attested`)
  return output(record(cwd, current, value, publish))
}

export function recordReviewFailure(cwd: string, issue: number, kind: Gate, reason: string, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  event(value, 'REVIEW_FAILED', `${kind}: ${reason}`)
  return output(record(cwd, current, value, publish))
}

export async function main(args = process.argv.slice(2)): Promise<void> {
  const parsed = parseArgs(args)
  const { command, values, flags } = parsed
  const cwd = resolve(values.cwd ?? process.cwd())
  const publish = !flags.has('no-github')
  const issue = values.issue ? issueNumber(values.issue) : undefined
  if (command === 'claim') {
    console.log(JSON.stringify(claim(cwd, required(values, 'issue'), required(values, 'branch'), publish), null, 2))
    return
  }
  if (!issue) throw new Error('--issue is required')
  if (command === 'spec' || command === 'plan') {
    console.log(JSON.stringify(setArtifact(cwd, issue, command, required(values, 'file'), publish), null, 2))
    return
  }
  if (command === 'review') {
    console.log(
      JSON.stringify(
        review(cwd, issue, required(values, 'gate') as Gate, required(values, 'brief'), publish, values.result),
        null,
        2,
      ),
    )
    return
  }
  if (command === 'review-failed') {
    console.log(
      JSON.stringify(
        recordReviewFailure(cwd, issue, required(values, 'gate') as Gate, required(values, 'reason'), publish),
        null,
        2,
      ),
    )
    return
  }
  if (command === 'checks') {
    console.log(JSON.stringify(recordGate(cwd, issue, required(values, 'command'), publish), null, 2))
    return
  }
  if (command === 'live') {
    console.log(JSON.stringify(recordLive(cwd, issue, required(values, 'evidence'), publish), null, 2))
    return
  }
  const state = output(load(cwd, issue))
  if (command === 'history') {
    console.log(JSON.stringify(state.events, null, 2))
    return
  }
  if (command === 'status') {
    console.log(JSON.stringify(state, null, 2))
    return
  }
  throw new Error(`unknown command: ${command}`)
}

if (import.meta.url === `file://${process.argv[1]}`)
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : error)
    process.exitCode = 1
  })
