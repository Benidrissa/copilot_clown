# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Copilot Clown provides `=USEAI()` and `=USEAI.SINGLE()` custom Excel functions — a BYOK alternative to Microsoft's `=COPILOT()`. It supports Anthropic Claude, OpenAI, and Google Gemini as LLM providers with in-memory response caching.

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
- `CopilotClown/Functions/UseAiFunction.cs` — `[ExcelFunction]` USEAI (spill) and USEAI.SINGLE (single cell), 8 params each (4 prompt + 4 context). Orchestrates file resolve → cache → rate limit (with wait time + alternative model suggestion) → file upload → API (with system prompt + generation params) → markdown strip → format (1D column, 2D table, or single cell). Contains `LoadingObservable` (IExcelObservable for "Loading..." indicator), `StripMarkdown()`, `FormatResponse()`, and `FindAlternativeModel()`.

**Services:**
- `CopilotClown/Services/LlmProvider.cs` — `ILlmProvider` interface (with `AppSettings`-aware `CompleteAsync` overload) + `ProviderFactory` (dictionary lookup for 3 providers)
- `CopilotClown/Services/ClaudeProvider.cs` — Anthropic Messages API. Supports system prompt (top-level `system` field), temperature, top_p. Default `max_tokens: 8192`. Supports multimodal content (text + images). File upload via Files API (`/v1/files`, `anthropic-beta: files-api-2025-04-14`). Prompt caching (`cache_control: ephemeral`) on attachment blocks.
- `CopilotClown/Services/OpenAIProvider.cs` — OpenAI Chat Completions API. Supports system prompt (`role: system` message), temperature, top_p (omitted for reasoning models o1/o3/o4). GPT-5.x uses `max_completion_tokens`; older models use `max_tokens`. Supports multimodal content (text + images). File upload via Files API (`/v1/files`, `purpose: user_data`).
- `CopilotClown/Services/GeminiProvider.cs` — Google Gemini Generate Content API. Supports system prompt (`systemInstruction` field), temperature, topP, maxOutputTokens. Supports multimodal content (text + images via `inlineData`/`fileData`). File upload via Gemini Files API (`media.upload`).
- `CopilotClown/Services/CacheService.cs` — `MemoryCache` with SHA-256 keys. Prompt-only cache keys (shared across providers/models). Fingerprints prompts >2048 chars. Thread-safe hit/miss counters.
- `CopilotClown/Services/FileResolver.cs` — Detects `{...}` file/folder/URL references in prompts, delegates extraction to `ContentExtractor`, returns `ResolvedPrompt`. Recursive folder support (max 50 files).
- `CopilotClown/Services/ContentExtractor.cs` — Extracts text from `.docx` (OpenXml), `.pdf` (PdfPig + Tesseract OCR for scanned pages), plain text files. Images read as base64. URL download via `HttpClient`.
- `CopilotClown/Services/ContentCache.cs` — Separate `MemoryCache` for extracted file content. Local files cached until cleared (key includes `LastWriteTimeUtc`). URLs cached with configurable TTL.
- `CopilotClown/Services/FileUploadCache.cs` — Maps `(provider, sourcePath, lastModified)` → remote `file_id`. Avoids re-uploading files to provider APIs. Local files infinite TTL (key includes `LastWriteTimeUtc`). URLs 24h TTL. Thread-safe counters.
- `CopilotClown/Services/SettingsService.cs` — JSON settings in `%APPDATA%\CopilotClown\settings.json`, API keys DPAPI-encrypted in `keys.dat`. In-memory cache with 5s TTL to avoid disk I/O per cell.
- `CopilotClown/Services/PromptBuilder.cs` — Static `Build(object[] args)`. Handles `string`, `double`, `bool`, `object[,]` ranges, `ExcelMissing`/`ExcelEmpty`. Uses `StringBuilder`.
- `CopilotClown/Services/RateLimiter.cs` — Sliding-window rate limiter (default 500 calls / 10 min). Per-provider instances (Anthropic, OpenAI, and Google rate-limited independently). Exposes `TimeUntilNextSlot`, `UsagePercentage`, `IsNearLimit`, `IsLimited`, `FormatWaitTime()` for the Rate Limit Tracker UI.
- `CopilotClown/Services/JsonHelper.cs` — Thin wrapper around `JavaScriptSerializer` with dot-path navigation (`GetValue`, `GetString`, `GetInt`).

**UI:**
- `CopilotClown/UI/RibbonController.cs` — Adds "AI Settings" and "Convert to Values" buttons to Home ribbon tab via `ExcelRibbon`. Also registers `[ExcelCommand]` `ShowAISettings` as Alt+F8 fallback.
- `CopilotClown/UI/SettingsForm.cs` — WinForms dialog with 7 tabs: API Keys (save/test per provider incl. Google), Model (3-provider radio + model dropdown with context window/pricing info), System Prompt (persistent persona text box), Parameters (temperature slider, max tokens, top_p), Cache (enable/disable, TTL dropdown, stats, clear), Rate Limits (per-provider status dashboard with auto-refresh, model availability matrix, reset counters), Tools (convert USEAI formulas to values for sharing).

**Models:**
- `CopilotClown/Models/Models.cs` — `ProviderName` enum (Anthropic, OpenAI, Google), `ModelInfo`, `CompletionResponse`, `AppSettings` (defaults: OpenAI/gpt-5.2, cache 24h, rate limit 500/10min, temperature 1.0, maxTokens 8192, topP 1.0, systemPrompt ""), `ModelRegistry` (9 Claude models, 20 OpenAI models, 4 Gemini models).
- `CopilotClown/Models/ContentModels.cs` — `AttachmentType` enum, `Attachment` (type, content, mime, source, `RemoteFileId`, `RawBytes`), `ResolvedPrompt` (clean text + attachments list).

## Docs

- `docs/SRS.md` — Software Requirements Specification
