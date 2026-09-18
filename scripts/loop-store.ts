import { execFileSync } from 'node:child_process'
import { randomUUID } from 'node:crypto'
import { unlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

export interface StateResult<T> {
  oid: string
  value: T
}

const ZERO_OID = '0'.repeat(40)

function cleanEnvironment(): NodeJS.ProcessEnv {
  return Object.fromEntries(Object.entries(process.env).filter(([key]) => !key.startsWith('GIT_')))
}

function git(cwd: string, args: string[], input?: string): string {
  return execFileSync('git', args, {
    cwd,
    encoding: 'utf8',
    input,
    env: cleanEnvironment(),
    stdio: ['pipe', 'pipe', 'pipe'],
  }).trim()
}

function refFor(issue: number): string {
  if (!Number.isSafeInteger(issue) || issue < 1) throw new Error(`invalid issue number: ${issue}`)
  return `refs/heads/loop-state/issue-${issue}`
}

function stable(value: unknown): string {
  if (value === undefined) throw new Error('cannot serialize undefined value into loop state')
  if (value === null || typeof value !== 'object') return JSON.stringify(value)
  if (Array.isArray(value)) return `[${value.map(stable).join(',')}]`
  const record = value as Record<string, unknown>
  const keys = Object.keys(record).filter((key) => record[key] !== undefined)
  return `{${keys
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stable(record[key])}`)
    .join(',')}}`
}

function readRemoteOid(cwd: string, issue: number, remote: string): string | null {
  const lines = git(cwd, ['ls-remote', remote, refFor(issue)])
  const line = lines.split('\n').find(Boolean)
  return line ? line.split('\t', 1)[0] : null
}

function readCommit<T>(cwd: string, oid: string): T {
  return JSON.parse(git(cwd, ['show', `${oid}:state.json`])) as T
}

function commitState<T>(cwd: string, value: T, parent?: string): string {
  const path = join(tmpdir(), `engineering-loop-${randomUUID()}.json`)
  try {
    writeFileSync(path, `${stable(value)}\n`, 'utf8')
    const blob = git(cwd, ['hash-object', '-w', path])
    const tree = git(cwd, ['mktree'], `100644 blob ${blob}\tstate.json\n`)
    const args = [
      '-c',
      'user.name=engineering-loop',
      '-c',
      'user.email=engineering-loop@localhost',
      'commit-tree',
      tree,
    ]
    if (parent) args.push('-p', parent)
    args.push('-m', 'engineering-loop state')
    return git(cwd, args)
  } finally {
    try {
      unlinkSync(path)
    } catch {
      /* best effort cleanup */
    }
  }
}

export function readState<T>(cwd: string, issue: number, remote = 'origin'): StateResult<T> | null {
  const oid = readRemoteOid(cwd, issue, remote)
  if (!oid) return null
  git(cwd, ['fetch', '--no-tags', remote, refFor(issue)])
  return { oid, value: readCommit<T>(cwd, oid) }
}

export function createState<T>(cwd: string, issue: number, value: T, remote = 'origin'): StateResult<T> {
  const ref = refFor(issue)
  const oid = commitState(cwd, value)
  try {
    git(cwd, ['push', '--no-verify', `--force-with-lease=${ref}:${ZERO_OID}`, remote, `${oid}:${ref}`])
  } catch (error) {
    throw new Error(
      `exclusive claim failed for issue #${issue}: ${error instanceof Error ? error.message : String(error)}`,
    )
  }
  const actual = readRemoteOid(cwd, issue, remote)
  if (actual !== oid) throw new Error(`exclusive claim lost for issue #${issue}`)
  return { oid, value }
}

export function updateState<T>(
  cwd: string,
  issue: number,
  expectedOid: string,
  value: T,
  remote = 'origin',
): StateResult<T> {
  const current = readRemoteOid(cwd, issue, remote)
  if (current !== expectedOid)
    throw new Error(`state conflict for issue #${issue}: expected ${expectedOid}, found ${current ?? 'missing'}`)
  const oid = commitState(cwd, value, expectedOid)
  const ref = refFor(issue)
  try {
    git(cwd, ['push', '--no-verify', `--force-with-lease=${ref}:${expectedOid}`, remote, `${oid}:${ref}`])
  } catch (error) {
    throw new Error(`state update lost for issue #${issue}: ${error instanceof Error ? error.message : String(error)}`)
  }
  const actual = readRemoteOid(cwd, issue, remote)
  if (actual !== oid) throw new Error(`state update lost for issue #${issue}`)
  return { oid, value }
}
