# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Copilot Clown provides an `=USEAI()` custom Excel function — a BYOK alternative to Microsoft's `=COPILOT()`. It supports both Anthropic Claude and OpenAI as LLM providers with response caching. The repo is a monorepo with two autonomous implementations:

| Version | Directory | Platform | Technology | Deployment |
|---------|-----------|----------|------------|------------|
| **Web** | `web/` | Windows, Mac, Web | Office.js, TypeScript, React | GitHub Pages (static host) |
| **.NET** | `dotnet/` | Windows only | Excel-DNA, C# .NET 8, WinForms | Self-contained .xll file |

## Web Add-in (`web/`)

```bash
cd web
npm install              # Install dependencies
npm run build            # Production build (webpack, localhost URLs)
npm run build:prod       # Production build with correct URLs via scripts/build.js
npm run dev              # Start webpack dev server (https://localhost:3000)
npm run start            # Sideload add-in into Excel
npm run validate         # Validate dist/manifest.xml
npm run lint             # ESLint check
```

**Static hosting build:**
```bash
node scripts/build.js https://username.github.io/copilot_clown
```

`manifest.xml` uses `{{BASE_URL}}` placeholders replaced at build time. Push to `main` triggers GitHub Actions deploy to Pages.

**Architecture:** Office.js shared runtime — custom function and task pane share a single JS context. Services instantiated in `functions.ts`, exposed via `window.aillmServices`.

**Key files:**
- `src/functions/functions.ts` — `=USEAI()` registration via `CustomFunctions.associate()`
- `src/services/providers/` — `LLMProvider` interface, `ClaudeProvider`, `OpenAIProvider`
- `src/services/CacheService.ts` — SHA-256 localStorage cache with TTL and LRU eviction
- `src/services/SettingsService.ts` — localStorage settings, API keys Base64-encoded
- `src/taskpane/` — React settings UI (API keys, model selector, cache manager)

**CORS note:** Claude provider sends `anthropic-dangerous-direct-browser-access: true` header for direct browser API calls.

## .NET Add-in (`dotnet/`)

```bash
cd dotnet
dotnet restore           # Restore NuGet packages
dotnet build             # Build (outputs .xll to bin/)
dotnet build -c Release  # Release build
```

**Output:** `CopilotClown/bin/{Debug|Release}/net8.0-windows/CopilotClown-AddIn64.xll` — drag into Excel or add via File > Options > Add-ins.

**Architecture:** Excel-DNA .xll add-in. Single self-contained file, no server needed. Uses `ExcelAsyncUtil.Run()` for non-blocking API calls from cells.

**Key files:**
- `CopilotClown/Functions/UseAiFunction.cs` — `[ExcelFunction]` USEAI with 8 params (4 prompt + 4 context). Orchestrates cache → rate limit → API → spill.
- `CopilotClown/Services/ClaudeProvider.cs` / `OpenAIProvider.cs` — `ILlmProvider` implementations
- `CopilotClown/Services/CacheService.cs` — `MemoryCache` with SHA-256 keys
- `CopilotClown/Services/SettingsService.cs` — JSON settings in `%APPDATA%\CopilotClown\`, API keys encrypted via DPAPI
- `CopilotClown/UI/RibbonController.cs` — Adds "AI Settings" button to Home ribbon tab
- `CopilotClown/UI/SettingsForm.cs` — WinForms dialog (tabs: API Keys, Model, Cache)
- `CopilotClown/Models/Models.cs` — `ModelRegistry`, `ProviderName` enum, records

## Shared

- `docs/SRS.md` — Software Requirements Specification (covers both versions)
- `.github/workflows/deploy.yml` — GitHub Actions: builds and deploys `web/` to GitHub Pages
