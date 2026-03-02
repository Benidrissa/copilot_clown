# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Copilot Clown is a Microsoft Excel Web Add-in (Office.js) that provides an `=USEAI()` custom function — a BYOK alternative to Microsoft's `=COPILOT()` function. It supports both Anthropic Claude and OpenAI as LLM providers with localStorage-based response caching.

## Build & Run Commands

```bash
npm install              # Install dependencies
npm run build            # Production build (webpack)
npm run build:dev        # Development build
npm run dev              # Start webpack dev server (https://localhost:3000)
npm run start            # Sideload add-in into Excel for testing
npm run stop             # Stop the sideloaded add-in
npm run validate         # Validate manifest.xml
npm run lint             # ESLint check
```

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

`manifest.xml` — XML format Office Add-in manifest. The function namespace is `COPILOTCLOWN`, so the full qualified function name in Excel is `=COPILOTCLOWN.USEAI()`. Dev server URL: `https://localhost:3000`.

## Anthropic Browser CORS

The Claude provider sends `anthropic-dangerous-direct-browser-access: true` header to allow direct browser-to-API calls. This is required because the add-in has no backend proxy.
