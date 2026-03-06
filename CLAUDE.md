# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Copilot Clown provides `=USEAI()` and `=USEAI.SINGLE()` custom Excel functions — a BYOK alternative to Microsoft's `=COPILOT()`. It supports both Anthropic Claude and OpenAI as LLM providers with in-memory response caching.

| Platform | Technology | Deployment |
|----------|------------|------------|
| Windows only | Excel-DNA, C# .NET 4.8, WinForms | Self-contained .xll file |

## Build & Run (`dotnet/`)

```bash
# dotnet SDK is installed at ~/.dotnet — add to PATH first
export PATH="$HOME/.dotnet:$PATH"

cd dotnet
dotnet restore           # Restore NuGet packages
dotnet build             # Build (outputs .xll to bin/)
dotnet build -c Release  # Release build
```

**Output:** `CopilotClown/bin/{Debug|Release}/net48/CopilotClown-AddIn64.xll` — drag into Excel or add via File > Options > Add-ins.

**Installer:** `installer/build-installer.ps1` runs `dotnet build -c Release` then compiles an Inno Setup installer (`CopilotClown.iss`) → `installer/Output/CopilotClownSetup.exe`.

**Architecture:** Excel-DNA .xll add-in. Single self-contained file (all deps packed), no server needed. Uses `ExcelAsyncUtil.Observe()` with `IExcelObservable` for non-blocking API calls from cells (shows "Loading..." during processing). Forces TLS 1.2/1.3. JSON via built-in `JavaScriptSerializer` (zero NuGet JSON deps).

## Key files

**Functions:**
- `CopilotClown/Functions/UseAiFunction.cs` — `[ExcelFunction]` USEAI (spill) and USEAI.SINGLE (single cell), 8 params each (4 prompt + 4 context). Orchestrates file resolve → cache → rate limit → file upload → API → markdown strip → format (1D column, 2D table, or single cell). Contains `LoadingObservable` (IExcelObservable for "Loading..." indicator), `StripMarkdown()`, and `FormatResponse()`.

**Services:**
- `CopilotClown/Services/LlmProvider.cs` — `ILlmProvider` interface + `ProviderFactory` (dictionary lookup)
- `CopilotClown/Services/ClaudeProvider.cs` — Anthropic Messages API. No system prompt. Default `max_tokens: 8192`. Supports multimodal content (text + images). File upload via Files API (`/v1/files`, `anthropic-beta: files-api-2025-04-14`). Prompt caching (`cache_control: ephemeral`) on attachment blocks.
- `CopilotClown/Services/OpenAIProvider.cs` — OpenAI Chat Completions API. No system prompt. GPT-5.x uses `max_completion_tokens`; older models use `max_tokens`. Supports multimodal content (text + images). File upload via Files API (`/v1/files`, `purpose: user_data`).
- `CopilotClown/Services/CacheService.cs` — `MemoryCache` with SHA-256 keys. Fingerprints prompts >2048 chars. Thread-safe hit/miss counters.
- `CopilotClown/Services/FileResolver.cs` — Detects `{...}` file/folder/URL references in prompts, delegates extraction to `ContentExtractor`, returns `ResolvedPrompt`. Recursive folder support (max 50 files).
- `CopilotClown/Services/ContentExtractor.cs` — Extracts text from `.docx` (OpenXml), `.pdf` (PdfPig + Tesseract OCR for scanned pages), plain text files. Images read as base64. URL download via `HttpClient`.
- `CopilotClown/Services/ContentCache.cs` — Separate `MemoryCache` for extracted file content. Local files cached until cleared (key includes `LastWriteTimeUtc`). URLs cached with configurable TTL.
- `CopilotClown/Services/FileUploadCache.cs` — Maps `(provider, sourcePath, lastModified)` → remote `file_id`. Avoids re-uploading files to provider APIs. Local files infinite TTL (key includes `LastWriteTimeUtc`). URLs 24h TTL. Thread-safe counters.
- `CopilotClown/Services/SettingsService.cs` — JSON settings in `%APPDATA%\CopilotClown\settings.json`, API keys DPAPI-encrypted in `keys.dat`. In-memory cache with 5s TTL to avoid disk I/O per cell.
- `CopilotClown/Services/PromptBuilder.cs` — Static `Build(object[] args)`. Handles `string`, `double`, `bool`, `object[,]` ranges, `ExcelMissing`/`ExcelEmpty`. Uses `StringBuilder`.
- `CopilotClown/Services/RateLimiter.cs` — Sliding-window rate limiter (default 500 calls / 10 min).
- `CopilotClown/Services/JsonHelper.cs` — Thin wrapper around `JavaScriptSerializer` with dot-path navigation (`GetValue`, `GetString`, `GetInt`).

**UI:**
- `CopilotClown/UI/RibbonController.cs` — Adds "AI Settings" and "Convert to Values" buttons to Home ribbon tab via `ExcelRibbon`. Also registers `[ExcelCommand]` `ShowAISettings` as Alt+F8 fallback.
- `CopilotClown/UI/SettingsForm.cs` — WinForms dialog with 4 tabs: API Keys (save/test per provider), Model (provider radio + model dropdown with context window/pricing info), Cache (enable/disable, TTL dropdown, stats, clear), Tools (convert USEAI formulas to values for sharing).

**Models:**
- `CopilotClown/Models/Models.cs` — `ProviderName` enum, `ModelInfo`, `CompletionResponse`, `AppSettings` (defaults: OpenAI/gpt-5.2, cache 24h, rate limit 500/10min), `ModelRegistry` (9 Claude models, 20 OpenAI models).
- `CopilotClown/Models/ContentModels.cs` — `AttachmentType` enum, `Attachment` (type, content, mime, source, `RemoteFileId`, `RawBytes`), `ResolvedPrompt` (clean text + attachments list).

## Docs

- `docs/SRS.md` — Software Requirements Specification
