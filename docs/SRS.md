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

- The function is **asynchronous** via `ExcelAsyncUtil.Run()` — Excel shows `#BUSY!` while waiting
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

- The user selects the active provider and model from the task pane
- The selection persists across sessions via localStorage
- All `=USEAI()` calls use the currently active provider/model
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

### 3.5 Rate Limiting

- Client-side rate limiter: configurable limit (default 500 calls per 10 minutes)
- Uses a sliding window counter stored in memory
- When limit is exceeded, `=USEAI()` returns error: `"Rate limit exceeded. Wait and try again."`
- Rate limit counter resets independently of Excel recalculation

---

## 4. Non-Functional Requirements

### 4.1 Performance

| Metric | Target |
|--------|--------|
| Cache lookup time | < 5 ms |
| UI responsiveness | Task pane loads in < 1 second |
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
│  │  =USEAI()    │    │       Task Pane (React)       │    │
│  │  Custom Fn   │    │  ┌────────────────────────┐   │    │
│  │              │    │  │  API Key Settings       │   │    │
│  └──────┬───────┘    │  │  Model Selector         │   │    │
│         │            │  │  Cache Manager           │   │    │
│         │            │  └────────────────────────┘   │    │
│         │            └──────────────┬────────────────┘    │
│         │    Shared Runtime (JS)    │                     │
│  ┌──────▼───────────────────────────▼──────────────┐     │
│  │                                                  │     │
│  │  ┌──────────────┐  ┌────────────────────────┐   │     │
│  │  │ PromptBuilder│  │   SettingsService      │   │     │
│  │  └──────┬───────┘  │   (localStorage)       │   │     │
│  │         │          └────────────────────────┘   │     │
│  │  ┌──────▼───────────────────────────────────┐   │     │
│  │  │           CacheService                    │   │     │
│  │  │    (SHA-256 hash → localStorage)          │   │     │
│  │  └──────┬───────────────────────────────────┘   │     │
│  │         │ (cache miss)                          │     │
│  │  ┌──────▼───────────────────────────────────┐   │     │
│  │  │        LLM Provider (abstract)            │   │     │
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
2. Excel calls the registered custom function handler
3. **PromptBuilder** resolves cell references and constructs the full prompt string
4. **CacheService** computes SHA-256 hash of `{ provider, model, prompt }` and checks localStorage
5. **Cache hit:** Return stored response immediately
6. **Cache miss:** **RateLimiter** checks call budget → **LLMProvider** sends API request
7. Response is stored in cache and returned to the cell
8. If response contains multiple items, it is returned as a 2D array (spill)

### 5.3 Shared Runtime Architecture

The add-in uses an Office.js **Shared Runtime** so that:
- Custom functions and the task pane share the same JavaScript execution context
- Settings changed in the task pane (API key, model) are immediately available to custom functions via `window` globals
- `localStorage` is accessible from both custom functions and the task pane
- No need for cross-runtime messaging or `OfficeRuntime.storage`

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
  "max_tokens": 4096,
  "messages": [
    {
      "role": "user",
      "content": "<constructed_prompt>"
    }
  ],
  "system": "You are a helpful assistant embedded in Microsoft Excel. Respond concisely and directly. When asked to produce a list, return one item per line with no numbering or bullets unless requested. Do not add explanations unless asked."
}
```

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
  "max_tokens": 4096,
  "messages": [
    {
      "role": "system",
      "content": "You are a helpful assistant embedded in Microsoft Excel. Respond concisely and directly. When asked to produce a list, return one item per line with no numbering or bullets unless requested. Do not add explanations unless asked."
    },
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

### 8.5 Cache Management (Settings Dialog)

- **Stats display:** Total entries, hits, misses, hit rate
- **Clear cache:** Trims `MemoryCache` (removes all entries)
- **TTL input:** Adjust cache duration
- **Enable/disable toggle:** Bypass cache without clearing it

---

## Appendix A: Function Examples

```excel
' Simple prompt
=USEAI("List 5 popular programming languages")

' With cell context
=USEAI("Summarize this feedback", A2:A20)

' Classification with categories
=USEAI("Classify", B2:B100, "as one of these sentiments:", D1:D3)

' Product description from specs
=USEAI("Write a product description based on:", B2:B8)

' Translation
=USEAI("Translate to French:", A2)
```

## Appendix B: Settings File Location

| File | Path | Purpose |
|------|------|---------|
| `settings.json` | `%APPDATA%\CopilotClown\settings.json` | Active provider, model, cache/rate-limit config |
| `keys.dat` | `%APPDATA%\CopilotClown\keys.dat` | API keys (DPAPI encrypted, JSON with Base64 values) |
