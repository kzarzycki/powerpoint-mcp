#!/usr/bin/env node

import { execFileSync, spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
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
export type Gate = 'spec' | 'plan' | 'branch' | 'checks'
export type ReviewGate = 'spec' | 'plan' | 'branch'

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
  artifacts: Partial<Record<'spec' | 'plan', { path: string }>>
  implementation?: { head: string }
  merge?: { pr: string; commit: string }
  live?: { status: 'evidence' | 'waived'; detail: string; at: string }
  reviews: Array<{
    gate: ReviewGate
    verdict: 'APPROVE' | 'REVISE'
    reviewer: string
    findings: string[]
    at: string
    source: 'omp' | 'file'
  }>
  gates: Array<{ gate: 'checks'; command: string; exitCode: number; output: string; at: string }>
  events: LoopEvent[]
  updatedAt: string
}

const MAX_ATTEMPTS = 3
const GATE_TIMEOUT_MS = 20 * 60_000
const EMPTY_ATTEMPTS: Record<Gate, number> = { spec: 0, plan: 0, branch: 0, checks: 0 }
const KNOWN_VALUE_KEYS = new Set([
  'cwd',
  'issue',
  'branch',
  'file',
  'pr',
  'gate',
  'brief',
  'result',
  'model',
  'command',
  'evidence',
  'waive',
])
const KNOWN_FLAG_KEYS = new Set(['no-github'])

function now(): string {
  return new Date().toISOString()
}
function session(): string {
  const value = process.env.AGENT_SESSION
  if (!value) throw new Error('AGENT_SESSION is required for loop mutations')
  return value
}
function cleanEnvironment(): NodeJS.ProcessEnv {
  const removed = new Set(['GIT_DIR', 'GIT_WORK_TREE', 'GIT_INDEX_FILE', 'GIT_COMMON_DIR', 'GIT_PREFIX'])
  return Object.fromEntries(Object.entries(process.env).filter(([key]) => !removed.has(key)))
}
function git(cwd: string, args: string[]): string {
  return execFileSync('git', args, { cwd, encoding: 'utf8', env: cleanEnvironment() }).trim()
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
function output(result: StateResult<LoopState>): LoopState {
  return result.value
}
function load(cwd: string, issue: number): StateResult<LoopState> {
  const state = readState<LoopState>(cwd, issue)
  if (!state) throw new Error(`issue #${issue} has no claim state`)
  return state
}
function update(cwd: string, state: StateResult<LoopState>, value: LoopState): StateResult<LoopState> {
  if (session() !== value.owner) throw new Error(`issue #${value.issue} is owned by ${value.owner}`)
  value.updatedAt = now()
  return updateState(cwd, value.issue, state.oid, value)
}
function event(value: LoopState, kind: string, detail: string, attempt?: number): LoopState {
  value.events.push({ at: now(), kind, phase: value.phase, detail, ...(attempt === undefined ? {} : { attempt }) })
  return value
}
function comment(issue: string, body: string): void {
  const result = spawnSync('gh', ['issue', 'comment', issue, '--body', body], { encoding: 'utf8' })
  if (result.status !== 0) console.error(`GitHub issue comment failed: ${result.stderr || result.stdout}`)
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
function requireImplementationHead(value: LoopState, cwd: string, action: string): string {
  if (!value.implementation) throw new Error(`issue #${value.issue} has no recorded implementation head`)
  const head = git(cwd, ['rev-parse', 'HEAD'])
  if (head !== value.implementation.head)
    throw new Error(
      `worktree HEAD ${head} has moved past the recorded implementation head ${value.implementation.head}; run 'implemented' again before ${action}`,
    )
  return head
}
function nextAttempt(value: LoopState, name: Gate): number {
  const next = value.attempts[name] + 1
  if (next > MAX_ATTEMPTS) throw new Error(`${name} gate is exhausted; issue is ${value.phase}`)
  return next
}
function parseArgs(args: string[]): { command: string; values: Record<string, string>; flags: Set<string> } {
  const [command = 'status', ...rest] = args
  const values: Record<string, string> = {}
  const flags = new Set<string>()
  for (let index = 0; index < rest.length; index += 1) {
    const token = rest[index]
    if (!token.startsWith('--')) throw new Error(`unexpected argument: ${token}`)
    const key = token.slice(2)
    if (KNOWN_VALUE_KEYS.has(key)) {
      const next = rest[index + 1]
      if (next === undefined || next.startsWith('--')) throw new Error(`--${key} requires a value`)
      values[key] = next
      index += 1
    } else if (KNOWN_FLAG_KEYS.has(key)) {
      flags.add(key)
    } else {
      throw new Error(`unknown option: --${key}`)
    }
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
  const worktree = resolve(dirname(repo), `powerpoint-mcp--issue-${issue}`)
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
    const winner = readState<LoopState>(repo, issue)
    if (winner && publish)
      comment(url, `engineering-loop: claim rejected — ${value.owner} lost to ${winner.value.owner}`)
    throw error
  }
  try {
    git(repo, ['fetch', '--no-tags', 'origin', 'main'])
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
  value.artifacts[kind] = { path: artifactPath }
  event(value, `${kind.toUpperCase()}_WRITTEN`, `${kind}=${artifactPath}`)
  return output(record(cwd, current, value, publish))
}
export function markImplemented(cwd: string, issue: number, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  const reopenable: Phase[] = ['PLAN_APPROVED', 'IMPLEMENTED', 'BRANCH_APPROVED', 'GATES_GREEN']
  if (!reopenable.includes(value.phase))
    throw new Error(`issue #${issue} is ${value.phase}; expected one of ${reopenable.join(', ')}`)
  const head = git(value.worktree, ['rev-parse', 'HEAD'])
  const wasApproved = value.phase === 'BRANCH_APPROVED' || value.phase === 'GATES_GREEN'
  const reopened = wasApproved && value.implementation?.head !== head
  const fromPhase = value.phase
  value.implementation = { head }
  if (!wasApproved || reopened) {
    value.phase = 'IMPLEMENTED'
    if (reopened) value.live = undefined
  }
  event(value, 'IMPLEMENTED', `branch head=${head}${reopened ? `; reopened for re-review after ${fromPhase}` : ''}`)
  return output(record(cwd, current, value, publish))
}
export function markMerged(cwd: string, issue: number, pr: string, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  requirePhase(value, 'GATES_GREEN')
  const head = requireImplementationHead(value, value.worktree, 'merging')
  if (!value.live)
    throw new Error(
      `issue #${issue} has no recorded live evidence or waiver; run 'live --evidence <path>' or 'live --waive <reason>' first`,
    )
  value.merge = { pr, commit: head }
  value.phase = 'MERGED'
  event(value, 'MERGED', `pr=${pr} commit=${head}`)
  return output(record(cwd, current, value, publish))
}
export function review(
  cwd: string,
  issue: number,
  kind: ReviewGate,
  briefPath: string | undefined,
  publish = true,
  resultPath?: string,
  model?: string,
): LoopState {
  if (!briefPath && !resultPath) throw new Error('--brief is required unless --result is provided')
  const current = load(cwd, issue)
  const value = output(current)
  const expected: Record<ReviewGate, Phase> = {
    spec: 'SPEC',
    plan: 'PLAN',
    branch: 'IMPLEMENTED',
  }
  requirePhase(value, expected[kind])
  if (kind === 'branch') requireImplementationHead(value, value.worktree, 'reviewing')
  const attempt = nextAttempt(value, kind)
  let result: ReviewResult
  try {
    if (resultPath) {
      const raw = JSON.parse(readFileSync(resolve(cwd, resultPath), 'utf8')) as Record<string, unknown>
      if (typeof raw.reviewer !== 'string') throw new Error('recorded review has no reviewer')
      result = {
        ...parseReviewOutput(JSON.stringify({ verdict: raw.verdict, findings: raw.findings })),
        reviewer: raw.reviewer,
      }
    } else
      result = runReview({ cwd: value.worktree, brief: readFileSync(resolve(cwd, briefPath as string), 'utf8'), model })
  } catch (error) {
    const failed = load(cwd, issue)
    const failedValue = output(failed)
    event(failedValue, 'REVIEW_FAILED', `${kind}: ${error instanceof Error ? error.message : String(error)}`)
    record(cwd, failed, failedValue, publish)
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
    source: resultPath ? 'file' : 'omp',
  })
  if (result.verdict === 'APPROVE') {
    const next: Record<ReviewGate, Phase> = {
      spec: 'SPEC_APPROVED',
      plan: 'PLAN_APPROVED',
      branch: 'BRANCH_APPROVED',
    }
    reviewed.phase = next[kind]
    event(reviewed, 'REVIEW_APPROVED', `${kind} approved by ${result.reviewer}`, attempt)
  } else {
    reviewed.attempts[kind] = attempt
    event(reviewed, 'REVIEW_REJECTED', `${kind} rejected by ${result.reviewer}: ${result.findings.join('; ')}`, attempt)
    if (attempt >= MAX_ATTEMPTS) {
      reviewed.phase = 'PARKED'
      event(reviewed, 'PARKED', `${kind} rejected three times`, attempt)
    }
  }
  return output(record(cwd, latest, reviewed, publish))
}
function gateCommand(command: string): { bin: string; args: string[] } {
  const override = process.env.LOOP_GATE_RUNNER
  const prefix = override === undefined ? ['mise', 'exec', 'node@24.18.0', '--'] : override.split(' ').filter(Boolean)
  const [bin, ...prefixArgs] = [...prefix, 'sh']
  return { bin, args: [...prefixArgs, '-lc', command] }
}
export function recordGate(cwd: string, issue: number, command: string, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  requirePhase(value, 'BRANCH_APPROVED')
  requireImplementationHead(value, value.worktree, 'checks')
  const { bin, args } = gateCommand(command)
  const result = spawnSync(bin, args, {
    cwd: value.worktree,
    encoding: 'utf8',
    maxBuffer: 2 * 1024 * 1024,
    timeout: GATE_TIMEOUT_MS,
  })
  const text = `${result.stdout ?? ''}${result.stderr ?? ''}${result.error ? `\n${result.error.message}` : ''}`.slice(
    -12000,
  )
  if (result.error || result.signal) {
    const detail = result.error?.message ?? `terminated by ${result.signal}`
    event(value, 'GATE_INFRA_FAILED', `checks: ${detail}`)
    record(cwd, current, value, publish)
    throw new Error(`checks command failed to run: ${detail}`)
  }
  const attempt = nextAttempt(value, 'checks')
  value.gates.push({ gate: 'checks', command, exitCode: result.status ?? 1, output: text, at: now() })
  if (result.status === 0) {
    value.phase = 'GATES_GREEN'
    event(value, 'GATE_PASSED', `checks attempt ${attempt}: exit 0; ${command}`, attempt)
  } else {
    value.attempts.checks = attempt
    event(value, 'GATE_FAILED', `checks attempt ${attempt}: exit ${result.status ?? 1}; ${command}`, attempt)
    if (attempt >= MAX_ATTEMPTS) {
      value.phase = 'PARKED'
      event(value, 'PARKED', 'checks rejected three times', attempt)
    }
  }
  return output(record(cwd, current, value, publish))
}
export function recordLive(cwd: string, issue: number, evidencePath: string, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  requirePhase(value, 'GATES_GREEN')
  const path = resolve(cwd, evidencePath)
  if (!existsSync(path)) throw new Error(`live evidence does not exist: ${path}`)
  value.live = { status: 'evidence', detail: path, at: now() }
  event(value, 'LIVE_EVIDENCE', `operator evidence=${path}; live Office prerequisite must be attested`)
  return output(record(cwd, current, value, publish))
}
export function waiveLive(cwd: string, issue: number, reason: string, publish = true): LoopState {
  const current = load(cwd, issue)
  const value = output(current)
  requirePhase(value, 'GATES_GREEN')
  if (!reason.trim()) throw new Error('a live gate waiver requires a reason')
  value.live = { status: 'waived', detail: reason, at: now() }
  event(value, 'LIVE_WAIVED', `operator waived the live gate: ${reason}`)
  return output(record(cwd, current, value, publish))
}

export async function main(args = process.argv.slice(2)): Promise<void> {
  const { command, values, flags } = parseArgs(args)
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
  if (command === 'implemented') {
    console.log(JSON.stringify(markImplemented(cwd, issue, publish), null, 2))
    return
  }
  if (command === 'merged') {
    console.log(JSON.stringify(markMerged(cwd, issue, required(values, 'pr'), publish), null, 2))
    return
  }
  if (command === 'review') {
    const gate = required(values, 'gate')
    if (gate !== 'spec' && gate !== 'plan' && gate !== 'branch')
      throw new Error(`--gate must be spec, plan or branch (got ${gate})`)
    console.log(JSON.stringify(review(cwd, issue, gate, values.brief, publish, values.result, values.model), null, 2))
    return
  }
  if (command === 'checks') {
    console.log(JSON.stringify(recordGate(cwd, issue, required(values, 'command'), publish), null, 2))
    return
  }
  if (command === 'live') {
    console.log(
      JSON.stringify(
        values.waive
          ? waiveLive(cwd, issue, values.waive, publish)
          : recordLive(cwd, issue, required(values, 'evidence'), publish),
        null,
        2,
      ),
    )
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
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : error)
    process.exitCode = 1
  })
