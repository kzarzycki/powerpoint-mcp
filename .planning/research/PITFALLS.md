# Domain Pitfalls: Open-Sourcing a Personal Node.js Tool

**Project:** PowerPoint Office.js Bridge
**Domain:** Open-sourcing a working personal project (~700 LOC) on GitHub
**Researched:** 2026-02-10

---

## Critical Pitfalls

Mistakes that would damage the project's reputation or create real security/legal risk.

### Pitfall 1: Personal Information Baked Into Git History

**What goes wrong:** The repo goes public with the author's home directory paths, email, and internal workflow tooling references preserved in every commit. Even if cleaned from current files, `git log -p` exposes them all. Someone runs `git log --all -p | grep /Users/` and finds full paths like `/Users/<you>/.claude/get-shit-done/workflows/...` and personal email `k.zarzycki@gmail.com` in every commit.

**Why it happens:** During personal development, hardcoded paths are convenient. Nobody thinks about git history until publish time.

**Consequences:**
- Personal paths leak the author's username and directory structure
- Internal workflow tool paths (`.claude/get-shit-done/`) are irrelevant to external users and look messy
- Email in commits is expected (and fine), but hardcoded paths in documentation/code commits look unprofessional

**Specific instances found in this project:**
- `CLAUDE.md` line 114: `"cwd": "/Users/<you>/dev/powerpoint-bridge"` (committed)
- `.planning/` files: 20+ references to `/Users/<you>/` paths across planning documents
- `.planning/phases/05-multi-session-support/05-RESEARCH.md`: contains `/Users/<you>/Library/Application Support/mkcert/rootCA.pem`
- Git author in all commits: `Krzysztof Zarzycki <k.zarzycki@gmail.com>` (this is normal and fine for open source)

**Prevention:**
- **Decision: fresh repo vs. history rewrite.** For a ~30 commit project, starting a fresh repo with a single "initial commit" is simpler and cleaner than using `git-filter-repo`. The planning history has no value to external contributors.
- If keeping history: use `git-filter-repo` to replace all `/Users/<you>/` with generic placeholders
- Before publishing, grep the entire repo (including `.planning/` if kept) for personal paths
- Replace hardcoded paths with `$PROJECT_ROOT` or relative paths in all docs

**Detection:** `git log --all -p | grep -i '/Users/'` and `grep -r '/Users/' .`

**Phase:** Must be addressed in the very first phase of open-source prep (repo hygiene), before anything goes public.

**Confidence:** HIGH -- verified by reading actual committed files and git history.

---

### Pitfall 2: Exposing Internal Planning Artifacts

**What goes wrong:** The `.planning/` directory (35+ files of GSD workflow artifacts -- milestone audits, phase research, plan documents, verification checklists, agent history) gets published. These are internal development scaffolding, not documentation. External users see agent configuration, execution plans with `@/Users/<you>/.claude/get-shit-done/` references, and "yolo mode" config -- none of which helps them use the tool.

**Why it happens:** The planning system is part of the working directory and tracked by git. It feels like "documentation" but is actually process artifacts.

**Consequences:**
- Confuses users who expect standard project docs (README, CONTRIBUTING)
- Exposes personal workflow tooling details
- Makes the repo look cluttered (35+ planning files vs ~6 source files)
- Agent history JSON and config with `"mode": "yolo"` creates odd first impression

**Prevention:**
- Add `.planning/` to `.gitignore` before the public repo is created
- If starting fresh repo: simply do not include `.planning/` in the initial commit
- If keeping history: `git-filter-repo --path .planning/ --invert-paths` to remove from all commits
- Keep `.planning/` locally for the author's workflow, just exclude from public repo

**Detection:** Check if `.planning/` is tracked: `git ls-files .planning/`

**Phase:** First phase (repo hygiene), same as personal path cleanup.

**Confidence:** HIGH -- verified by listing 35+ files in `.planning/` directory.

---

### Pitfall 3: No LICENSE File = Not Actually Open Source

**What goes wrong:** The repo goes public without a LICENSE file. Despite being on GitHub, the code has no legal permission for anyone to use, modify, or distribute it. GitHub's terms allow viewing/forking but not using the code. Some developers and all organizations will skip the project entirely.

**Why it happens:** Personal projects don't need licenses. The author forgets this is the legal foundation of open source.

**Consequences:**
- Legally, nobody can use the code (copyright defaults to all rights reserved)
- GitHub shows "No license" badge -- immediate credibility hit
- Organizations cannot adopt it (legal/compliance teams reject unlicensed code)
- Contributors won't submit PRs to unlicensed repos (their employer's IP policy)

**Prevention:**
- Add MIT LICENSE file (standard for dev tools, maximum adoption)
- Add `"license": "MIT"` to package.json (currently missing)
- Do this in the first phase, not as an afterthought

**Detection:** `ls LICENSE*` and check `license` field in package.json.

**Phase:** First phase (repo hygiene).

**Confidence:** HIGH -- verified package.json has no license field, no LICENSE file exists.

---

### Pitfall 4: Security Concern -- Arbitrary Code Execution via MCP

**What goes wrong:** The `execute_officejs` tool sends arbitrary JavaScript to the PowerPoint add-in for execution via `AsyncFunction` constructor (equivalent to `eval`). This is by design -- it is the core architecture. But without clear security documentation, the project gets flagged as insecure, or users do not understand the trust model.

**Why it happens:** The tool is designed for a trusted local environment (Claude Code talking to your own PowerPoint). But open-source users may not understand the threat model.

**Consequences:**
- Security-conscious users reject the project without reading further
- Someone files a "critical vulnerability" issue for the intended behavior
- If the bridge is ever exposed beyond localhost, arbitrary code execution is a real risk

**Specific code paths:**
- `addin/app.js` line 113: `var fn = new AsyncFunction('context', 'PowerPoint', code);` -- constructs and executes arbitrary code
- `server/index.ts` line 280: `execute_officejs` tool accepts raw code string
- Server binds to `localhost` only (good) but this is not explicitly documented as a security boundary

**Prevention:**
- Document the trust model clearly in README: "This tool runs on localhost. The MCP client (Claude Code) sends code to your local PowerPoint. Do not expose the bridge server to a network."
- Add a "Security" section to README explaining: localhost-only, trusted client model, no authentication (because localhost)
- Consider adding a SECURITY.md file for responsible disclosure
- The `localhost` binding is the security boundary -- make sure it stays that way and document it

**Detection:** Search for `AsyncFunction`, `eval`, `new Function` in codebase.

**Phase:** Documentation phase (README writing). Does not require code changes.

**Confidence:** HIGH -- verified by reading `addin/app.js` and `server/index.ts`.

---

## Moderate Pitfalls

Mistakes that create a bad first impression or cause setup failures for new users.

### Pitfall 5: The mkcert Setup Creates an Interactive Roadblock

**What goes wrong:** A new user clones the repo, runs `npm install`, tries `npm start`, and gets "TLS certificate files not found." They run `npm run setup-certs`, which calls `mkcert`, but first they need `mkcert -install` which prompts for their macOS Keychain password. The user has no idea why a PowerPoint tool needs Keychain access. They close the terminal and never come back.

**Why it happens:** WSS is mandatory (macOS WKWebView requirement), so TLS certs are non-negotiable. But `mkcert -install` is an interactive step that cannot be automated or scripted.

**Consequences:**
- First-run experience has an unexplained password prompt
- Users on CI/CD cannot run the setup
- Users without Homebrew need to install that first too
- The "Prerequisites: brew install mkcert node" instruction assumes Homebrew

**Prevention:**
- README must explain WHY certs are needed (WKWebView enforces WSS) before asking user to install mkcert
- Provide a step-by-step "First-time setup" section with clear expectations: "You will be prompted for your macOS password once to trust the local certificate authority"
- The `npm run setup-certs` script already exists but does not include `mkcert -install` -- document this as a separate prerequisite step
- Consider a setup script that checks for mkcert, checks if CA is installed, and guides the user through each step with clear messages
- For CI: document that tests should mock the WebSocket layer, not require real TLS

**Detection:** Try `npm start` in a fresh clone without certs. Does the error message guide the user?

**Phase:** Documentation phase and possibly a small code improvement (better error message in server startup).

**Confidence:** HIGH -- verified by reading `server/index.ts` lines 28-35 and `package.json` scripts.

---

### Pitfall 6: No README = Invisible Project

**What goes wrong:** Someone finds the repo via search or a link. They see the GitHub page with no README -- just a file listing of `addin/`, `server/`, `CLAUDE.md`, `RESEARCH.md`. They have no idea what this project does, who it is for, or how to use it. They leave.

**Why it happens:** The project was built for personal use with `CLAUDE.md` serving as the primary documentation (for AI agents, not humans).

**Consequences:**
- GitHub renders nothing on the repo landing page
- No architecture overview, no screenshots, no quick start
- CLAUDE.md is agent-facing documentation, not human-facing
- RESEARCH.md is research notes, not user docs
- The project's unique value proposition ("first live-editing MCP bridge for PowerPoint on macOS") is invisible

**Prevention:**
- Write README.md with: one-line description, architecture diagram, prerequisites, quick start, usage examples, limitations
- Move CLAUDE.md content that is human-relevant into README, keep CLAUDE.md as agent instructions only
- Consider a demo GIF showing Claude Code creating slides in real-time
- The architecture diagram from CLAUDE.md is excellent -- adapt it for README

**Detection:** Does `README.md` exist? Does the GitHub landing page explain the project in 10 seconds?

**Phase:** Core documentation phase.

**Confidence:** HIGH -- verified no README.md exists.

---

### Pitfall 7: Incomplete .gitignore Leaks IDE and OS Files

**What goes wrong:** The current `.gitignore` only has 3 lines: `node_modules/`, `certs/`, `*.pem`. A contributor on macOS commits `.DS_Store` files. Someone using VS Code commits `.vscode/settings.json` with personal preferences. The repo accumulates junk files.

**Why it happens:** Personal projects start with minimal `.gitignore` because only one person works on them.

**Specific gaps in current `.gitignore`:**
- No `.DS_Store` (macOS Finder metadata -- will appear in every directory)
- No `.vscode/`, `.idea/`, `*.swp` (editor configs)
- No `.env` (if anyone adds environment variables later)
- No `dist/`, `build/`, `*.tgz` (build artifacts)
- No `.planning/` (internal workflow files)
- No `.mcp.json` (local MCP configuration with localhost URL)

**Prevention:**
- Use a comprehensive Node.js + macOS `.gitignore` template
- Add `.planning/`, `.mcp.json`, and `.claude/` to `.gitignore`
- Check existing tracked files: `git ls-files` -- if `.mcp.json` is already tracked, `git rm --cached .mcp.json`

**Detection:** Compare `.gitignore` against GitHub's standard Node.js template.

**Phase:** First phase (repo hygiene).

**Confidence:** HIGH -- verified by reading `.gitignore` (3 lines).

---

### Pitfall 8: macOS-Only Without Saying So

**What goes wrong:** A Windows or Linux user clones the repo and spends 30 minutes trying to set it up before discovering it only works on macOS. The sideloading path (`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`) is macOS-specific. The WKWebView WSS requirement is macOS-specific. Nothing in the repo says "macOS only."

**Why it happens:** The author develops on macOS and the entire stack assumes it. This is not a bug -- Office.js add-in sideloading on macOS genuinely works differently than on Windows.

**Consequences:**
- Wasted time for non-macOS users
- Negative first impression ("doesn't work")
- Issues filed for Windows support that are actually platform limitations

**Prevention:**
- Put "macOS" in the repo description and first line of README
- Add a "Platform" or "Requirements" badge/section at the very top
- Document what would need to change for Windows (different sideloading path, WKWebView not an issue, WSS may not be required)
- Consider: is the architecture fundamentally macOS-only, or just the setup? (Answer: mostly setup -- the add-in/server architecture could work on Windows with different sideloading and potentially plain WS)

**Detection:** Does the README mention the platform requirement in the first 5 lines?

**Phase:** Documentation phase.

**Confidence:** HIGH -- verified by examining sideloading paths and WSS requirements.

---

### Pitfall 9: package.json Missing Open-Source Metadata

**What goes wrong:** The `package.json` looks like a private project. No `repository`, `homepage`, `bugs`, `keywords`, `engines`, `license`, or `description` fields. npm and GitHub tooling cannot link the package to the repo. `npm info` shows minimal metadata.

**Current package.json gaps:**
- No `description` field
- No `license` field
- No `repository` field
- No `homepage` field
- No `bugs` field
- No `engines` field (Node.js 24 is required for native TS execution -- this is critical)
- No `keywords` field
- `"private": true` should be set (to prevent accidental npm publish)

**Prevention:**
- Add all standard metadata fields
- `engines` is especially important: the server runs `.ts` files directly via Node 24's native TypeScript support -- this will fail silently or with confusing errors on Node 18/20
- Add `"private": true` since this is not an npm package

**Detection:** Run `npm pack --dry-run` and check what would be published.

**Phase:** Repo hygiene phase.

**Confidence:** HIGH -- verified by reading package.json.

---

### Pitfall 10: CLAUDE.md Contains Stale Architecture Description

**What goes wrong:** `CLAUDE.md` describes `MCP stdio` transport in the architecture diagram and MCP configuration, but the project actually uses HTTP transport on port 3001. The build order still says "Phase 2: add MCP server (stdio)". A contributor reads CLAUDE.md, gets confused by the mismatch with actual code.

**Specific stale content:**
- Architecture diagram line 10: `Claude Code <--MCP stdio-->` (should be HTTP)
- MCP Configuration section lines 108-118: shows stdio config with `"command": "node"` (actual config is HTTP on port 3001)
- Build order references stdio transport

**Why it happens:** CLAUDE.md was written at project start and not updated when the transport changed from stdio to HTTP in phase 05-02.

**Prevention:**
- Update CLAUDE.md architecture diagram and MCP configuration to reflect current HTTP transport
- Or, if CLAUDE.md is being slimmed down for open source, ensure the README has accurate architecture info
- Add a note about the architectural evolution if keeping historical context

**Detection:** Compare CLAUDE.md claims against actual `.mcp.json` and `server/index.ts` imports.

**Phase:** Documentation phase.

**Confidence:** HIGH -- verified by comparing CLAUDE.md lines 10/108-118 against `.mcp.json` and server imports.

---

## Minor Pitfalls

Mistakes that cause annoyance but are easily fixable.

### Pitfall 11: No CONTRIBUTING.md = Contributor Confusion

**What goes wrong:** Someone wants to contribute. They do not know: how to set up the dev environment, what the code style is, how to run tests (there are none yet), or what the PR process looks like.

**Prevention:**
- Add CONTRIBUTING.md with: dev setup, code style expectations, how to test, PR process
- Keep it short -- this is a small project, not a framework

**Phase:** Documentation phase.

**Confidence:** HIGH.

---

### Pitfall 12: No Tests = No Confidence for Contributors

**What goes wrong:** A contributor wants to refactor the WebSocket connection pool. There are no tests to verify they did not break anything. They either skip the contribution or submit untested changes.

**Prevention:**
- Add basic test coverage for: command protocol serialization, connection pool management, MCP tool parameter validation
- Do not over-test -- this is ~700 LOC. A handful of unit tests for the server logic is sufficient.
- The add-in (Office.js) is hard to test outside PowerPoint -- document this limitation

**Phase:** Testing phase (after core documentation).

**Confidence:** HIGH.

---

### Pitfall 13: Hardcoded Port Numbers Without Configuration

**What goes wrong:** A user already has something running on port 8443 or 3001. The server fails to start with `EADDRINUSE`. There is no way to change ports without editing source code.

**Current state:** Ports 8443 and 3001 are hardcoded in `server/index.ts` lines 18-19. The add-in also hardcodes `wss://localhost:8443` in `addin/app.js` line 34, and the manifest.xml references `https://localhost:8443` throughout.

**Prevention:**
- Support environment variables: `PORT=8443` and `MCP_PORT=3001` with current values as defaults
- Document the port configuration in README
- Note that changing the HTTPS port requires updating `manifest.xml` and re-sideloading (this is a real limitation worth documenting)

**Phase:** Could be deferred to post-initial-release. Document the limitation for now.

**Confidence:** HIGH -- verified by reading source code.

---

## Phase-Specific Warnings

| Phase Topic | Likely Pitfall | Mitigation |
|-------------|---------------|------------|
| Repo hygiene | Personal paths in git history (Pitfall 1) | Fresh repo or git-filter-repo |
| Repo hygiene | .planning/ exposure (Pitfall 2) | Exclude from public repo |
| Repo hygiene | Missing LICENSE (Pitfall 3) | MIT license, first commit |
| Repo hygiene | Incomplete .gitignore (Pitfall 7) | Use comprehensive template |
| Repo hygiene | package.json gaps (Pitfall 9) | Add all metadata fields |
| Documentation | Security concerns not addressed (Pitfall 4) | Trust model in README |
| Documentation | mkcert friction undocumented (Pitfall 5) | Step-by-step setup guide |
| Documentation | No README (Pitfall 6) | Write human-facing docs |
| Documentation | macOS-only not stated (Pitfall 8) | Platform badge, first line |
| Documentation | Stale CLAUDE.md (Pitfall 10) | Update or slim down |
| Documentation | No CONTRIBUTING.md (Pitfall 11) | Brief contributor guide |
| Testing | No tests (Pitfall 12) | Basic server unit tests |
| Post-release | Hardcoded ports (Pitfall 13) | Environment variable support |

## Summary of Severity

- **Must fix before publish:** Pitfalls 1, 2, 3, 6, 7 (repo hygiene)
- **Should fix before publish:** Pitfalls 4, 5, 8, 9, 10 (documentation quality)
- **Nice to have for publish:** Pitfalls 11, 12, 13 (contributor experience)

## Sources

- Git history analysis: `git log --all -p` on local repository
- Codebase inspection: all source files in `server/`, `addin/`, root config files
- [GitHub Docs: Removing sensitive data from a repository](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository)
- [GitHub opensource.guide](https://opensource.guide/)
- [10up Open Source Best Practices](https://10up.github.io/Open-Source-Best-Practices/)
- [npm package.json documentation](https://docs.npmjs.com/cli/v7/configuring-npm/package-json/)
- [mkcert project](https://github.com/FiloSottile/mkcert) -- setup friction analysis
- [OWASP: Direct Dynamic Code Evaluation](https://owasp.org/www-community/attacks/Direct_Dynamic_Code_Evaluation_Eval%20Injection)
