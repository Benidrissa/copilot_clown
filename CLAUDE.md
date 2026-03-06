# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Copilot Clown provides an `=USEAI()` custom Excel function — a BYOK alternative to Microsoft's `=COPILOT()`. It supports both Anthropic Claude and OpenAI as LLM providers with response caching.

| Platform | Technology | Deployment |
|----------|------------|------------|
| Windows only | Excel-DNA, C# .NET 4.8, WinForms | Self-contained .xll file |

## Build & Run (`dotnet/`)

```bash
cd dotnet
dotnet restore           # Restore NuGet packages
dotnet build             # Build (outputs .xll to bin/)
dotnet build -c Release  # Release build
```

**Output:** `CopilotClown/bin/{Debug|Release}/net48/CopilotClown-AddIn64.xll` — drag into Excel or add via File > Options > Add-ins.

**Architecture:** Excel-DNA .xll add-in. Single self-contained file, no server needed. Uses `ExcelAsyncUtil.Run()` for non-blocking API calls from cells.

**Key files:**
- `CopilotClown/Functions/UseAiFunction.cs` — `[ExcelFunction]` USEAI with 8 params (4 prompt + 4 context). Orchestrates cache → rate limit → API → spill.
- `CopilotClown/Services/ClaudeProvider.cs` / `OpenAIProvider.cs` — `ILlmProvider` implementations
- `CopilotClown/Services/CacheService.cs` — `MemoryCache` with SHA-256 keys
- `CopilotClown/Services/SettingsService.cs` — JSON settings in `%APPDATA%\CopilotClown\`, API keys encrypted via DPAPI
- `CopilotClown/UI/RibbonController.cs` — Adds "AI Settings" button to Home ribbon tab
- `CopilotClown/UI/SettingsForm.cs` — WinForms dialog (tabs: API Keys, Model, Cache)
- `CopilotClown/Models/Models.cs` — `ModelRegistry`, `ProviderName` enum, records

## Docs

- `docs/SRS.md` — Software Requirements Specification
