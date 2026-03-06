# Software Requirements Specification (SRS)

## Copilot Clown — Excel AI Add-in

**Version:** 1.0
**Date:** 2026-03-02
**Status:** Draft

---

## Table of Contents

1. [Introduction](#1-introduction)
2. [Overall Description](#2-overall-description)
3. [Functional Requirements](#3-functional-requirements)
4. [Non-Functional Requirements](#4-non-functional-requirements)
5. [System Architecture](#5-system-architecture)
6. [API Specifications](#6-api-specifications)
7. [Error Handling](#7-error-handling)
8. [Caching Specification](#8-caching-specification)

---

## 1. Introduction

### 1.1 Purpose

This document specifies the requirements for **Copilot Clown**, a Microsoft Excel add-in (Excel-DNA .xll) that provides custom worksheet functions `=USEAI()` and `=USEAI.SINGLE()` enabling users to call AI language models (Claude by Anthropic and OpenAI GPT models) directly from Excel cells. The add-in is a BYOK (Bring Your Own Key) alternative to Microsoft's native `=COPILOT()` function.

### 1.2 Scope

The add-in provides:
- Custom Excel functions `=USEAI()` (spills multi-line responses) and `=USEAI.SINGLE()` (single cell)
- Support for Anthropic Claude and OpenAI as LLM providers
- User-managed API keys with provider and model switching
- In-memory response caching with SHA-256 keys to reduce API costs
- A WinForms settings dialog for configuration and cache management
- Markdown stripping for clean Excel output

### 1.3 Definitions and Acronyms

| Term | Definition |
|------|-----------|
| **BYOK** | Bring Your Own Key — user supplies their own API credentials |
| **Excel-DNA** | Open-source framework for building high-performance Excel add-ins in .NET |
| **XLL** | Excel add-in file format loaded directly by Excel |
| **Dynamic Array / Spill** | An Excel feature where a formula result fills multiple cells automatically |
| **TTL** | Time-To-Live — duration a cached entry remains valid |
| **LLM** | Large Language Model |
| **DPAPI** | Windows Data Protection API — encrypts data to the current user account |

### 1.4 References

- Excel-DNA: https://excel-dna.net/
- Anthropic Claude API: https://docs.anthropic.com/en/docs/api-reference
- OpenAI Chat Completions API: https://platform.openai.com/docs/api-reference/chat
- Microsoft COPILOT function documentation (source specification)

---

## 2. Overall Description

### 2.1 Product Perspective

Copilot Clown is a standalone Excel-DNA .xll add-in for Windows. It is distributed as a single self-contained file — users drag it into Excel or register it via File > Options > Add-ins. API calls are made directly from the user's machine to the respective AI provider endpoints. No backend server is required.

### 2.2 User Characteristics

Target users are Excel power users who:
- Have API accounts with Anthropic and/or OpenAI
- Want AI-assisted text generation, classification, summarization, and data enrichment within Excel
- Want to control costs and model selection (unlike Microsoft's locked-in COPILOT function)

### 2.3 Constraints

- **Platform:** Windows only (Excel-DNA requires Windows Excel)
- **Runtime:** .NET Framework 4.8
- **Network:** Requires internet connectivity to reach AI provider APIs
- **TLS:** Forces TLS 1.2/1.3 (required by both Anthropic and OpenAI APIs)
- **Storage:** Settings stored in `%APPDATA%\CopilotClown\`; cache in-memory (`MemoryCache`)

### 2.4 Assumptions

- Users have valid API keys for at least one supported provider
- Users understand that AI responses are non-deterministic and may vary
- Users are running Windows with Excel (2010 or later with dynamic array support for spill)

---

## 3. Functional Requirements

### 3.1 Custom Function: `=USEAI()`

#### 3.1.1 Syntax

```
=USEAI(prompt_part1, [context1], [prompt_part2], [context2], [prompt_part3], [context3], [prompt_part4], [context4])
=USEAI.SINGLE(prompt_part1, [context1], [prompt_part2], [context2], [prompt_part3], [context3], [prompt_part4], [context4])
```

`USEAI` spills multi-line responses into separate rows. `USEAI.SINGLE` returns the full response in a single cell with line breaks.

#### 3.1.2 Arguments

| Argument | Required | Type | Description |
|----------|----------|------|-------------|
| `prompt_part1` | Yes | String | Text describing the task or question |
| `context1` | No | Range or String | Cell reference or range providing data context |
| `prompt_part2` | No | String | Additional prompt text |
| `context2` | No | Range or String | Additional context |
| `prompt_part3` | No | String | Additional prompt text |
| `context3` | No | Range or String | Additional context |
| `prompt_part4` | No | String | Additional prompt text |
| `context4` | No | Range or String | Additional context |

The function accepts up to 8 arguments (4 prompt + 4 context pairs). Excel-DNA requires fixed parameter counts.

#### 3.1.3 Prompt Construction

Arguments are processed sequentially to build the full prompt:
- **String arguments** are appended as-is
- **Cell references** (single cell) are resolved to their text value
- **Range references** are resolved to a comma-separated list of cell values (or newline-separated for column ranges)

**Example:**
```
=USEAI("Classify", B1:B10, "into one of the following categories", A1:A4)
```
Produces prompt: `"Classify [values from B1:B10] into one of the following categories [values from A1:A4]"`

#### 3.1.4 Return Values

- **Single value:** Returns a string to the calling cell
- **Multiple values:** Returns a 2D array (`string[][]`) that spills to adjacent cells as a Dynamic Array
- The function determines array vs. single return based on the AI response content (if it contains line-separated items or structured data, it spills)

#### 3.1.5 Behavior

- The function is **asynchronous** via `ExcelAsyncUtil.Observe()` with `IExcelObservable` — the cell shows `"Loading..."` while the API call is in progress, then updates to the final result
- Results are **non-deterministic** — recalculating may produce different output
- The function checks the in-memory cache before making an API call
- If no API key is configured, returns a descriptive error string
- Markdown formatting is automatically stripped from AI responses
- Tabular responses (pipe-delimited) are detected and spilled as 2D arrays with numeric parsing
- Cell values are truncated to Excel's 32,767 character limit

### 3.2 Provider Support

#### 3.2.1 Supported Providers

| Provider | API Endpoint | Auth Method |
|----------|-------------|-------------|
| **Anthropic Claude** | `https://api.anthropic.com/v1/messages` | `x-api-key` header |
| **OpenAI** | `https://api.openai.com/v1/chat/completions` | `Authorization: Bearer` header |

#### 3.2.2 Available Models

**Claude Models:**

| Model ID | Display Name |
|----------|-------------|
| `claude-opus-4-6` | Claude Opus 4.6 |
| `claude-opus-4-5-20251101` | Claude Opus 4.5 |
| `claude-opus-4-1-20250805` | Claude Opus 4.1 |
| `claude-opus-4-20250514` | Claude Opus 4 |
| `claude-sonnet-4-6` | Claude Sonnet 4.6 |
| `claude-sonnet-4-5-20250929` | Claude Sonnet 4.5 |
| `claude-sonnet-4-20250514` | Claude Sonnet 4 |
| `claude-haiku-4-5-20251001` | Claude Haiku 4.5 |
| `claude-3-haiku-20240307` | Claude 3 Haiku |

**OpenAI Models:**

| Model ID | Display Name |
|----------|-------------|
| `gpt-5.2` | GPT-5.2 |
| `gpt-5.1` | GPT-5.1 |
| `gpt-5` | GPT-5 |
| `gpt-5-mini` | GPT-5 Mini |
| `gpt-5-nano` | GPT-5 Nano |
| `gpt-5-pro` | GPT-5 Pro |
| `gpt-5.2-pro` | GPT-5.2 Pro |
| `gpt-4.1` | GPT-4.1 |
| `gpt-4.1-mini` | GPT-4.1 Mini |
| `gpt-4.1-nano` | GPT-4.1 Nano |
| `gpt-4o` | GPT-4o |
| `gpt-4o-mini` | GPT-4o Mini |
| `gpt-4-turbo` | GPT-4 Turbo |
| `o4-mini` | o4-mini |
| `o3` | o3 |
| `o3-mini` | o3-mini |
| `o1` | o1 |
| `o1-pro` | o1 Pro |
| `gpt-3.5-turbo` | GPT-3.5 Turbo |

The model list is hardcoded but designed for easy extension via a registry pattern.

#### 3.2.3 Provider Switching

- The user selects the active provider and model from the settings dialog (Home ribbon > "AI Settings")
- The selection persists across sessions via `%APPDATA%\CopilotClown\settings.json`
- All `=USEAI()` and `=USEAI.SINGLE()` calls use the currently active provider/model
- Changing provider/model does **not** automatically recalculate existing formulas

### 3.3 API Key Management

#### 3.3.1 Storage
- API keys are stored in `%APPDATA%\CopilotClown\keys.dat`
- Keys are encrypted per-user via Windows DPAPI (`ProtectedData.Protect` with `DataProtectionScope.CurrentUser`)
- Keys are cached in memory for 5 seconds to avoid disk I/O on every cell calculation

#### 3.3.2 Settings UI
- WinForms settings dialog accessible from the "AI Settings" button on the Home ribbon tab
- Text input fields for each provider's API key (masked)
- "Test" button makes a minimal API call to validate the key
- Visual indicator showing key validity status

### 3.4 Settings Dialog

The settings dialog is a WinForms form with tabbed sections:

#### 3.4.1 API Keys Tab
- Input fields for Claude and OpenAI API keys
- Save and test functionality per key

#### 3.4.2 Model Selection Tab
- Provider toggle (Claude / OpenAI)
- Model dropdown filtered by selected provider
- Shows model capabilities summary (context window, pricing tier)

#### 3.4.3 Cache Management Tab
- Display: total cached entries, hit rate
- "Clear Cache" button (trims MemoryCache)
- TTL configuration input (default: 24 hours)
- Toggle to enable/disable caching globally

#### 3.4.4 Tools Tab
- "Convert Selected Cells" button — replaces formulas in current selection with their computed values
- "Convert All USEAI Cells" button — finds and converts all USEAI formulas in the workbook
- Status label showing conversion result count

### 3.5 File and Folder Content Attachment

#### 3.5.1 Overview

Users can reference local files, folders, and web URLs directly within prompt text by enclosing paths in curly braces `{ }`. The add-in auto-detects these references, extracts content (text or image), and attaches it to the API request.

#### 3.5.2 Syntax

File and folder references are enclosed in curly braces within the prompt:

```
=USEAI("Summarize {C:\reports\Q4.pdf}")
=USEAI("Describe {C:\photos\chart.png}")
=USEAI("Analyze all docs in {C:\reports\quarterly\}")
=USEAI("Summarize {https://example.com/report.pdf}")
```

#### 3.5.3 Supported File Types

| Category | Extensions | Processing |
|----------|-----------|------------|
| **Documents** | `.docx` | Text extraction via OpenXml (paragraph text) |
| **PDF** | `.pdf` | Text extraction via PdfPig; OCR fallback (Tesseract) for scanned pages |
| **Images** | `.png`, `.jpg`, `.jpeg`, `.gif`, `.webp`, `.bmp` | Base64-encoded, sent as multimodal content block |
| **Plain text** | `.txt`, `.csv`, `.md`, `.json`, `.xml`, `.html`, `.htm`, `.log`, `.yaml`, `.yml` | Read as-is (UTF-8, max 1 MB) |

Unsupported file types are silently skipped.

#### 3.5.4 Folder Processing

- Folders are processed **recursively** (all subfolders included)
- Maximum **50 files** per folder reference (alphabetical order)
- Only supported file types are included; others are skipped
- Path replaced with `[attached: N files from foldername]`

#### 3.5.5 URL Support

- HTTP/HTTPS URLs are downloaded via `HttpClient`
- Content type determined from HTTP `Content-Type` header or URL extension
- Maximum download size: **50 MB**
- Processed identically to local files after download

#### 3.5.6 Multimodal API Content

When attachments include images, the API request uses a multimodal content array:

- **Anthropic Claude:** `content` becomes an array of `{"type":"text",...}` and `{"type":"image","source":{"type":"base64",...}}` blocks. When a `file_id` is available (see §3.6), file references are used instead: `{"type":"document","source":{"type":"file","file_id":"..."}}` or `{"type":"image","source":{"type":"file","file_id":"..."}}`. Prompt caching (`cache_control`) is applied to attachment blocks.
- **OpenAI:** `content` becomes an array of `{"type":"text",...}` and `{"type":"image_url","image_url":{"url":"data:mime;base64,..."}}` blocks. When a `file_id` is available, file references are used: `{"type":"file","file":{"file_id":"..."}}`.
- **Text-only attachments:** Document text is prepended to the prompt as `--- Content from filename ---` sections; no multimodal array needed

#### 3.5.7 File Content Caching

- Extracted file content is cached in a dedicated `MemoryCache` ("FileContent"), separate from API response cache
- **Local files:** Cached indefinitely (until user clears cache). Cache key includes `LastWriteTimeUtc`, so modified files are automatically re-extracted
- **URLs:** Cached with configurable TTL (default: 24 hours)
- File content cache can be cleared independently via the Settings dialog
- The API response cache key includes attachment metadata (source path + content length) so that file changes invalidate cached API responses

#### 3.5.8 OCR (Optical Character Recognition)

- PDF pages with no extractable text (or fewer than 10 characters) are treated as scanned
- Scanned pages are OCR'd using Tesseract engine (English language)
- Tesseract data files (`eng.traineddata`) are located via: assembly directory → `TESSDATA_PREFIX` environment variable → `Program Files\Tesseract-OCR\tessdata`
- If OCR is unavailable, a warning message is returned inline: `[OCR unavailable: ...]`

#### 3.5.9 Error Handling

File resolution is best-effort — errors never prevent the API call:

| Scenario | Behavior |
|----------|----------|
| File not found | `[Warning: file not found]` inline |
| Unsupported extension | Skipped silently |
| Extraction failure | `[Error: could not read file: reason]` inline |
| URL download failure | `[Error: could not download: reason]` inline |
| Image > 20 MB | Skipped with warning |
| Folder > 50 files | First 50 processed, warning added |
| OCR failure | `[OCR error: reason]` inline |

### 3.6 API File Upload Caching

#### 3.6.1 Overview

To avoid re-uploading the same file on every API call, files are uploaded once to each provider's Files API and the returned `file_id` is cached. Subsequent requests reference the `file_id` instead of re-sending content inline.

#### 3.6.2 Provider File Upload APIs

| Provider | Endpoint | Auth | Purpose Field | Returns |
|----------|----------|------|---------------|---------|
| **Anthropic Claude** | `POST https://api.anthropic.com/v1/files` | `x-api-key` + `anthropic-beta: files-api-2025-04-14` | — | `file_id` |
| **OpenAI** | `POST https://api.openai.com/v1/files` | `Authorization: Bearer` | `purpose: "user_data"` | `file_id` |

Both APIs accept multipart form-data uploads with the file bytes and filename.

#### 3.6.3 File Reference Formats

| Provider | Document Reference | Image Reference |
|----------|-------------------|-----------------|
| **Claude** | `{"type":"document","source":{"type":"file","file_id":"..."},"cache_control":{"type":"ephemeral"}}` | `{"type":"image","source":{"type":"file","file_id":"..."},"cache_control":{"type":"ephemeral"}}` |
| **OpenAI** | `{"type":"file","file":{"file_id":"..."}}` | `{"type":"file","file":{"file_id":"..."}}` |

#### 3.6.4 Claude Prompt Caching

Anthropic's prompt caching is enabled on all attachment content blocks via `"cache_control": {"type": "ephemeral"}`. This provides up to 90% cost reduction when the same files are referenced in subsequent API calls within the caching window.

#### 3.6.5 Upload Flow

1. Before calling the completion API, each attachment is checked against `FileUploadCache`
2. **Cache hit:** `file_id` is set on the attachment; no upload needed
3. **Cache miss:** File bytes are uploaded to the provider's Files API; returned `file_id` is cached
4. **Upload failure:** Falls back silently to inline content (base64 or text)
5. The provider's completion request uses `file_id` references where available

#### 3.6.6 File Upload Cache

| Aspect | Detail |
|--------|--------|
| **Cache name** | `"FileUpload"` (separate `MemoryCache` instance) |
| **Local file key** | `upload_{provider}_{normalized_path}_{LastWriteTimeUtc.Ticks}` |
| **URL key** | `upload_{provider}_url_{url}` |
| **Local file TTL** | Infinite (until user clears or file modified → different key) |
| **URL TTL** | 24 hours |
| **Value** | Remote `file_id` string |
| **Eviction** | Manual clear via Settings dialog ("Clear File Cache" button) |

### 3.7 Rate Limiting

- Client-side rate limiter: configurable limit (default 500 calls per 10 minutes)
- Uses a sliding window counter stored in memory
- When limit is exceeded, `=USEAI()` returns error: `"Rate limit exceeded. Wait and try again."`
- Rate limit counter resets independently of Excel recalculation

### 3.8 Convert to Values

#### 3.8.1 Overview

Users can convert `=USEAI()` and `=USEAI.SINGLE()` formula cells to plain text values, enabling workbook sharing with users who do not have the add-in installed.

#### 3.8.2 Convert Selected Cells

- Accessible via ribbon button "Convert to Values" on the Home tab, or via the Tools tab in Settings
- Replaces formulas in the current Excel selection with their computed values
- Operates on all formulas in the selection, not limited to USEAI

#### 3.8.3 Convert All USEAI Cells

- Accessible via the Tools tab in the Settings dialog
- Iterates all worksheets in the active workbook
- Finds cells containing USEAI formulas (case-insensitive match on formula text)
- Replaces each formula with its current computed value
- Reports the count of converted cells

---

## 4. Non-Functional Requirements

### 4.1 Performance

| Metric | Target |
|--------|--------|
| Cache lookup time | < 5 ms |
| UI responsiveness | Settings dialog opens in < 1 second |
| API call timeout | 30 seconds (configurable) |
| Function registration | < 500 ms on Excel startup |

### 4.2 Security

- API keys are encrypted with DPAPI and stored locally only (never transmitted to third parties)
- API calls go directly from the client to the AI provider endpoints
- No telemetry, analytics, or data collection
- Prompts and context data are sent only to the user-selected AI provider
- The add-in has no server-side component

### 4.3 Compatibility

| Platform | Support Level |
|----------|--------------|
| Excel for Windows (Microsoft 365) | Full |
| Excel for Windows (2010+) | Full (dynamic arrays require Excel 365/2021) |
| Excel for Mac | Not supported |
| Excel Online (web) | Not supported |

### 4.4 Usability

- Function autocomplete in the formula bar with parameter hints (via ExcelDna.IntelliSense)
- Descriptive error messages (not just `#VALUE!`)
- Settings dialog accessible from "AI Settings" button on the Home ribbon tab

---

## 5. System Architecture

### 5.1 High-Level Component Diagram

```
┌─────────────────────────────────────────────────────────┐
│                    Excel Workbook                         │
│                                                           │
│  ┌──────────────┐    ┌──────────────────────────────┐    │
│  │  =USEAI()    │    │  Settings Dialog (WinForms)   │    │
│  │  =USEAI.     │    │  ┌────────────────────────┐   │    │
│  │    SINGLE()  │    │  │  API Key Settings       │   │    │
│  └──────┬───────┘    │  │  Model Selector         │   │    │
│         │            │  │  Cache Manager           │   │    │
│         │            │  └────────────────────────┘   │    │
│         │            └──────────────┬────────────────┘    │
│         │  Excel-DNA .xll (.NET 4.8)│                     │
│  ┌──────▼───────────────────────────▼──────────────┐     │
│  │                                                  │     │
│  │  ┌──────────────┐  ┌────────────────────────┐   │     │
│  │  │ PromptBuilder│  │   SettingsService      │   │     │
│  │  └──────┬───────┘  │   (%APPDATA% + DPAPI)  │   │     │
│  │         │          └────────────────────────┘   │     │
│  │  ┌──────▼───────────────────────────────────┐   │     │
│  │  │           CacheService                    │   │     │
│  │  │    (SHA-256 hash → MemoryCache)           │   │     │
│  │  └──────┬───────────────────────────────────┘   │     │
│  │         │ (cache miss)                          │     │
│  │  ┌──────▼───────────────────────────────────┐   │     │
│  │  │        ILlmProvider (interface)           │   │     │
│  │  │  ┌───────────────┐ ┌──────────────────┐  │   │     │
│  │  │  │ClaudeProvider │ │ OpenAIProvider    │  │   │     │
│  │  │  └───────┬───────┘ └────────┬─────────┘  │   │     │
│  │  └──────────┼──────────────────┼─────────────┘   │     │
│  └─────────────┼──────────────────┼─────────────────┘     │
└────────────────┼──────────────────┼───────────────────────┘
                 │                  │
                 ▼                  ▼
     ┌───────────────────┐  ┌──────────────────┐
     │ Anthropic API     │  │  OpenAI API      │
     │ api.anthropic.com │  │  api.openai.com  │
     └───────────────────┘  └──────────────────┘
```

### 5.2 Data Flow

1. User enters `=USEAI("Summarize", A1:A10)` in a cell
2. Excel calls the registered `[ExcelFunction]` handler
3. **PromptBuilder** resolves cell references / ranges and constructs the full prompt string
4. **FileResolver** detects `{...}` references in the prompt, extracts file/URL content (checking **ContentCache** first), and returns a `ResolvedPrompt` with clean text + attachments
5. **CacheService** computes SHA-256 hash of `{ provider, model, prompt + attachment metadata }` and checks `MemoryCache`
6. **Cache hit:** Return cached response immediately (already markdown-stripped)
7. **Cache miss:** `ExcelAsyncUtil.Observe()` dispatches async work (cell shows `"Loading..."`) → **RateLimiter** checks call budget → **FileUploadCache** checks if attachments need uploading (uploads via Files API if not cached) → **ILlmProvider** sends HTTP request (using `file_id` references or inline content)
8. Response is markdown-stripped, cached, and formatted for Excel (single value or 2D array)
9. If `USEAI`: multi-line responses spill as rows; tabular (pipe-delimited) responses spill as 2D arrays with numeric parsing. If `USEAI.SINGLE`: full text returned in one cell with embedded line breaks.

### 5.3 Excel-DNA Architecture

The add-in uses **Excel-DNA** (.NET Framework 4.8):
- Functions are marked with `[ExcelFunction]` attributes and compiled into a `.xll` file
- `ExcelAsyncUtil.Observe()` with `IExcelObservable` enables non-blocking API calls — cells show `"Loading..."` during execution, then update to the final result
- The ".xll" packages all dependencies into a single file (`ExcelDnaPackCompressResources`)
- The ribbon is extended with "AI Settings" and "Convert to Values" buttons via `ExcelRibbon` / `CustomUI`
- A fallback `[ExcelCommand]` macro (Alt+F8 > `ShowAISettings`) is registered in case the ribbon doesn't load
- Settings and API keys persist to `%APPDATA%\CopilotClown\` on disk
- JSON serialization uses `JavaScriptSerializer` (built into .NET 4.8, zero NuGet dependencies)
- Cache uses `System.Runtime.Caching.MemoryCache` (in-process, not persisted across sessions)

---

## 6. API Specifications

### 6.1 Anthropic Claude — Messages API

**Endpoint:** `POST https://api.anthropic.com/v1/messages`

**Request Headers:**
```
x-api-key: <user_api_key>
anthropic-version: 2023-06-01
content-type: application/json
```

**Request Body:**
```json
{
  "model": "claude-sonnet-4-20250514",
  "max_tokens": 8192,
  "messages": [
    {
      "role": "user",
      "content": "<constructed_prompt>"
    }
  ]
}
```

Note: No `system` message is sent by the .NET provider. The `max_tokens` default is 8192.

**Response (extract):**
```json
{
  "content": [
    {
      "type": "text",
      "text": "<response_text>"
    }
  ]
}
```

### 6.2 OpenAI — Chat Completions API

**Endpoint:** `POST https://api.openai.com/v1/chat/completions`

**Request Headers:**
```
Authorization: Bearer <user_api_key>
Content-type: application/json
```

**Request Body:**
```json
{
  "model": "gpt-4.1-mini",
  "max_tokens": 8192,
  "messages": [
    {
      "role": "user",
      "content": "<constructed_prompt>"
    }
  ]
}
```

Note: GPT-5.x models use `max_completion_tokens` instead of `max_tokens`. No `system` message is sent by the .NET provider. The `max_tokens` default is 8192.

```json
{
  "model": "gpt-5.2",
  "max_completion_tokens": 8192,
  "messages": [
    {
      "role": "user",
      "content": "<constructed_prompt>"
    }
  ]
}
```

**Response (extract):**
```json
{
  "choices": [
    {
      "message": {
        "content": "<response_text>"
      }
    }
  ]
}
```

---

## 7. Error Handling

### 7.1 Error Types

| Condition | Returned Value | Excel Error |
|-----------|---------------|-------------|
| No API key configured | `"Error: No API key configured for <provider>. Click 'AI Settings' in the ribbon."` | — |
| Empty prompt | `"Error: Prompt cannot be empty."` | `#VALUE!` |
| Prompt too large (>100,000 chars) | `"Error: Prompt too large. Reduce prompt or context size."` | `#VALUE!` |
| Rate limit exceeded (client-side) | `"Error: Rate limit exceeded. Wait and try again."` | — |
| API rate limit (429) | `"Error: API rate limit reached. Wait and retry."` | — |
| API authentication error (401/403) | `"Error: Invalid API key for <provider>. Check your key in the settings."` | — |
| Network timeout | `"Error: Request timed out. Try a simpler prompt or check your connection."` | — |
| Network error | `"Error: Network error. Check your internet connection."` | — |
| API server error (5xx) | `"Error: <Provider> API is unavailable. Try again later."` | — |
| Unknown error | `"Error: Unexpected error occurred."` | — |

### 7.2 Error Strategy

Custom functions return descriptive error strings rather than Excel error codes where possible, because string errors are more informative to the user.

---

## 8. Caching Specification

### 8.1 Cache Key Generation

```
cacheKey = SHA-256( provider + "|" + model + "|" + promptFingerprint )
```

For prompts ≤ 2048 characters, the full prompt is hashed. For longer prompts, a fingerprint is used: `length + first 512 chars + last 512 chars`.

The hash is hex-encoded and prefixed: `aillm_<hex_hash>`

### 8.2 Cache Storage

The cache uses `System.Runtime.Caching.MemoryCache` (in-process). Cache entries are not persisted across Excel sessions.

### 8.3 Cache Behavior

| Scenario | Action |
|----------|--------|
| Cache hit (within TTL) | Return cached response; no API call |
| Cache hit (expired TTL) | Entry auto-evicted by MemoryCache; make new API call; cache result |
| Cache miss | Make API call; cache result |
| Cache disabled (user toggle) | Always make API call; do not store |

### 8.4 Cache Configuration

| Setting | Default | Range |
|---------|---------|-------|
| TTL | 24 hours (1440 minutes) | 1 minute – 30 days |
| Caching enabled | `true` | `true` / `false` |
| Max cache entries | 1000 | 100 – 10,000 |

### 8.5 File Content Cache

| Aspect | Detail |
|--------|--------|
| **Cache name** | `"FileContent"` (separate `MemoryCache` instance) |
| **Local file key** | `file_{normalized_path}_{LastWriteTimeUtc.Ticks}` |
| **URL key** | `url_{url}` |
| **Local file TTL** | Infinite (until user clears or file modified → different key) |
| **URL TTL** | Configurable (default: 24 hours) |
| **Eviction** | Manual clear via Settings dialog |

### 8.6 Cache Management (Settings Dialog)

- **Stats display:** Total entries, hits, misses, hit rate
- **Clear cache:** Trims `MemoryCache` (removes all entries)
- **TTL input:** Adjust cache duration
- **Enable/disable toggle:** Bypass cache without clearing it

---

## Appendix A: Function Examples

```excel
' Simple prompt (spills multi-line response into rows)
=USEAI("List 5 popular programming languages")

' With cell context
=USEAI("Summarize this feedback", A2:A20)

' Classification with categories
=USEAI("Classify", B2:B100, "as one of these sentiments:", D1:D3)

' Product description from specs
=USEAI("Write a product description based on:", B2:B8)

' Translation
=USEAI("Translate to French:", A2)

' Full response in one cell (use Wrap Text to see all lines)
=USEAI.SINGLE("Write a paragraph about:", A2)

' Single-cell summary
=USEAI.SINGLE("Summarize this data in 2 sentences:", B1:B50)

' File attachment — PDF document
=USEAI("Summarize {C:\reports\Q4-2025.pdf}")

' File attachment — Word document
=USEAI("Extract key findings from {C:\docs\analysis.docx}")

' File attachment — image (sent as multimodal content)
=USEAI("Describe what you see in {C:\photos\chart.png}")

' Folder attachment — all supported files recursively
=USEAI("Summarize all reports in {C:\reports\quarterly\}")

' URL attachment — remote document
=USEAI("Summarize {https://example.com/report.pdf}")

' Mixed — file + text context
=USEAI("Compare {C:\old-report.pdf} with this data:", A1:A20)
```

## Appendix B: Settings File Location

| File | Path | Purpose |
|------|------|---------|
| `settings.json` | `%APPDATA%\CopilotClown\settings.json` | Active provider, model, cache/rate-limit config |
| `keys.dat` | `%APPDATA%\CopilotClown\keys.dat` | API keys (DPAPI encrypted, JSON with Base64 values) |
