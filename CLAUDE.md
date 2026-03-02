# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Copilot Clown is a Microsoft Excel Web Add-in (Office.js) that provides an `=USEAI()` custom function — a BYOK alternative to Microsoft's `=COPILOT()` function. It supports both Anthropic Claude and OpenAI as LLM providers with localStorage-based response caching.

## Build & Run Commands

```bash
npm install              # Install dependencies
npm run build            # Production build (webpack, localhost URLs)
npm run build:dev        # Development build
npm run build:prod       # Production build with correct URLs via scripts/build.js
npm run dev              # Start webpack dev server (https://localhost:3000)
npm run start            # Sideload add-in into Excel for testing
npm run stop             # Stop the sideloaded add-in
npm run validate         # Validate dist/manifest.xml
npm run lint             # ESLint check
```

### Production Build for Static Hosting

```bash
# Build for GitHub Pages (or any static host)
node scripts/build.js https://username.github.io/copilot_clown

# Or via env var
BASE_URL=https://example.com npm run build:prod
```

The `manifest.xml` source uses `{{BASE_URL}}` placeholders. The build script (`scripts/build.js`) or webpack replaces them with the actual hosting URL. Output goes to `dist/`.

## Architecture

**Shared Runtime** — The add-in uses an Office.js shared runtime, meaning the custom function and task pane run in the same JavaScript context. Services are instantiated once in `functions.ts` and exposed via `window.aillmServices` for the task pane to consume.

### Key Components

- **`src/functions/functions.ts`** — Registers `=USEAI()` custom function via `CustomFunctions.associate()`. Orchestrates: prompt building → cache check → rate limit → API call → cache store → response parsing (single value or spill array).
- **`src/services/providers/`** — `LLMProvider` interface with `ClaudeProvider` (Anthropic Messages API) and `OpenAIProvider` (Chat Completions API). Provider factory in `LLMProvider.ts` via `getProvider()`.
- **`src/services/CacheService.ts`** — SHA-256 hash-keyed localStorage cache with TTL, LRU eviction, and hit/miss stats. Keys prefixed `aillm_cache_`.
- **`src/services/SettingsService.ts`** — Reads/writes all user config to localStorage. API keys stored Base64-encoded with prefix `aillm_key_`.
- **`src/services/PromptBuilder.ts`** — Constructs a single prompt string from interleaved text and cell range arguments.
- **`src/services/RateLimiter.ts`** — In-memory sliding window rate limiter (default: 100 calls / 10 min).
- **`src/taskpane/`** — React task pane with three sections: API key management, model selection, cache management.

### Type System

All shared types in `src/types/index.ts` — includes model registry (`CLAUDE_MODELS`, `OPENAI_MODELS`), settings defaults, provider/cache interfaces.

## Manifest

`manifest.xml` — XML template with `{{BASE_URL}}` placeholders. The function namespace is `COPILOTCLOWN`, so the full qualified function name in Excel is `=COPILOTCLOWN.USEAI()`. The final manifest with resolved URLs is output to `dist/manifest.xml` during build.

## Deployment

Push to `main` → GitHub Actions (`.github/workflows/deploy.yml`) builds and deploys to GitHub Pages automatically. The workflow infers the Pages URL from the repo name. To sideload in Excel, use `dist/manifest.xml` after building.

## Anthropic Browser CORS

The Claude provider sends `anthropic-dangerous-direct-browser-access: true` header to allow direct browser-to-API calls. This is required because the add-in has no backend proxy.
