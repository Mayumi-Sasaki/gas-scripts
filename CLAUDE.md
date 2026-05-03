# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script (GAS) project. Scripts are deployed and run on Google's servers via the Apps Script runtime.

## Git Operation Rules

**Every code change must be committed and pushed to GitHub immediately after the change is made.**

```powershell
git add .
git commit -m "<type>: <summary>"
git push origin main
```

Commit message types: `feat`, `fix`, `refactor`, `docs`, `chore`

Do not batch multiple unrelated changes into a single commit. Each logical change gets its own commit and push.

## Development Setup

### Clasp (CLI for Apps Script)

```powershell
npm install -g @google/clasp
clasp login
clasp push   # upload local files to Apps Script
clasp pull   # download from Apps Script to local
clasp open   # open the script in the browser editor
```

### Project Structure

| File | Purpose |
|------|---------|
| `.clasp.json` | Clasp config (scriptId, rootDir) |
| `appsscript.json` | Apps Script manifest (scopes, runtime version) |
| `*.js` / `*.ts` | Script source files |

### Pushing Changes to Apps Script

After committing and pushing to GitHub, also push to Apps Script:

```powershell
clasp push
```

If TypeScript is used, clasp transpiles automatically — no separate build step needed.
