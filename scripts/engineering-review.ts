#!/usr/bin/env node

import { spawnSync } from 'node:child_process'
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

export type ReviewVerdict = 'APPROVE' | 'REVISE'

export interface ReviewResult {
  verdict: ReviewVerdict
  findings: string[]
  reviewer: string
}

export interface ReviewOutput {
  verdict: ReviewVerdict
  findings: string[]
}

export interface RunReviewOptions {
  cwd: string
  brief: string
  model?: string
  timeoutMs?: number
}

const DEFAULT_TIMEOUT_MS = 300_000

/** Parse the reviewer's JSON response, rejecting anything outside the review contract. */
export function parseReviewOutput(raw: string): ReviewOutput {
  if (typeof raw !== 'string' || raw.trim() === '') {
    throw new Error('Review output is empty')
  }

  let value: unknown
  try {
    value = JSON.parse(raw)
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    throw new Error(`Review output is not valid JSON: ${detail}`)
  }

  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('Review output must be a JSON object')
  }

  const keys = Object.keys(value)
  if (keys.length !== 2 || !keys.includes('verdict') || !keys.includes('findings')) {
    throw new Error('Review output must contain only verdict and findings')
  }

  const record = value as Record<string, unknown>
  if (record.verdict !== 'APPROVE' && record.verdict !== 'REVISE') {
    throw new Error("Review verdict must be 'APPROVE' or 'REVISE'")
  }
  const findings = record.findings
  if (!Array.isArray(findings) || !findings.every((finding): finding is string => typeof finding === 'string')) {
    throw new Error('Review findings must be an array of strings')
  }

  return {
    verdict: record.verdict,
    findings,
  }
}

function reviewPrompt(brief: string): string {
  return [
    'Review the engineering brief below as an independent reviewer.',
    'Assess whether the proposed work is ready to proceed and identify any concrete issues.',
    '',
    'Engineering brief:',
    brief,
    '',
    'Return ONLY one JSON object with exactly these keys: {"verdict":"APPROVE"|"REVISE","findings":[string]}.',
    'The verdict must be exactly APPROVE or REVISE. findings must be an array of strings. Do not use markdown, code fences, or additional keys.',
  ].join('\n')
}

export function parseOmpOutput(raw: string): ReviewOutput {
  try {
    return parseReviewOutput(raw)
  } catch (directError) {
    const texts: string[] = []
    for (const line of raw.split(/\r?\n/).reverse()) {
      try {
        const record = JSON.parse(line) as {
          type?: string
          message?: { content?: Array<{ type?: string; text?: string }> }
        }
        if (record.type !== 'turn_end') continue
        for (const item of record.message?.content ?? []) if (item.type === 'text' && item.text) texts.push(item.text)
        break
      } catch {
        // Continue looking for the final JSON event.
      }
    }
    if (texts.length === 0) throw directError
    return parseReviewOutput(texts.reverse().join(''))
  }
}

/** Run the local omp reviewer and parse its fail-closed JSON response. */
export function runReview(options: RunReviewOptions): ReviewResult {
  const { cwd, brief, model, timeoutMs = DEFAULT_TIMEOUT_MS } = options
  if (typeof cwd !== 'string' || cwd.trim() === '') throw new Error('Review cwd is required')
  if (typeof brief !== 'string' || brief.trim() === '') throw new Error('Review brief is required')
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new Error('Review timeout must be positive')

  const omp = process.env.OMP_BIN?.trim() || 'omp'
  const args = [
    '--mode',
    'json',
    '-p',
    reviewPrompt(brief),
    '--no-session',
    '--no-extensions',
    '--no-skills',
    '--no-rules',
    '--tools',
    'read,grep,glob',
    '--cwd',
    cwd,
    '--thinking',
    'high',
  ]
  if (model !== undefined) {
    if (model.trim() === '') throw new Error('Review model cannot be empty')
    args.push('--model', model)
  }

  const completed = spawnSync(omp, args, {
    cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: timeoutMs,
    maxBuffer: 4 * 1024 * 1024,
  })
  if (completed.error) throw new Error(`Review command failed: ${completed.error.message}`)
  if (completed.status !== 0) {
    const detail = completed.stderr.trim()
    throw new Error(`Review command exited with status ${completed.status}${detail ? `: ${detail}` : ''}`)
  }
  if (completed.signal) throw new Error(`Review command terminated by ${completed.signal}`)

  const parsed = parseOmpOutput(completed.stdout)
  return {
    ...parsed,
    reviewer: `omp:${model ?? 'default'}`,
  }
}

interface CliOptions {
  cwd: string
  briefFile: string
  output: string
  model?: string
  timeoutMs?: number
}

function parseCliArgs(argv: string[]): CliOptions {
  const values: Record<string, string | undefined> = {}
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index]
    if (
      flag !== '--cwd' &&
      flag !== '--brief-file' &&
      flag !== '--output' &&
      flag !== '--model' &&
      flag !== '--timeout-ms'
    ) {
      throw new Error(`Unknown argument: ${flag}`)
    }
    const value = argv[index + 1]
    if (value === undefined || value.startsWith('--')) throw new Error(`Missing value for ${flag}`)
    if (Object.hasOwn(values, flag)) throw new Error(`Duplicate argument: ${flag}`)
    values[flag] = value
    index += 1
  }

  const cwd = values['--cwd']
  const briefFile = values['--brief-file']
  const output = values['--output']
  if (cwd === undefined || briefFile === undefined || output === undefined) {
    throw new Error(
      'Usage: engineering-review.ts --cwd <dir> --brief-file <file> --output <file> [--model <name>] [--timeout-ms <n>]',
    )
  }
  const timeoutMs = values['--timeout-ms'] === undefined ? undefined : Number(values['--timeout-ms'])
  if (timeoutMs !== undefined && (!Number.isFinite(timeoutMs) || timeoutMs <= 0))
    throw new Error('--timeout-ms must be positive')
  return { cwd, briefFile, output, model: values['--model'], timeoutMs }
}

function main(): void {
  try {
    const options = parseCliArgs(process.argv.slice(2))
    const result = runReview({
      cwd: options.cwd,
      brief: readFileSync(options.briefFile, 'utf8'),
      model: options.model,
      timeoutMs: options.timeoutMs,
    })
    writeFileSync(options.output, `${JSON.stringify(result)}\n`, 'utf8')
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exitCode = 1
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main()
