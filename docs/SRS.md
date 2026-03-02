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

This document specifies the requirements for **Copilot Clown**, a Microsoft Excel Web Add-in that provides a custom worksheet function `=USEAI()` enabling users to call AI language models (Claude by Anthropic and OpenAI GPT models) directly from Excel cells. The add-in is a BYOK (Bring Your Own Key) alternative to Microsoft's native `=COPILOT()` function.

### 1.2 Scope

The add-in provides:
- A custom Excel function `=USEAI()` with syntax compatible with Microsoft's `=COPILOT()` function
- Support for Anthropic Claude and OpenAI as LLM providers
- User-managed API keys with provider and model switching
- Local response caching to reduce API costs
- A task pane UI for configuration and cache management

### 1.3 Definitions and Acronyms

| Term | Definition |
|------|-----------|
| **BYOK** | Bring Your Own Key — user supplies their own API credentials |
| **Custom Function** | A user-defined Excel formula registered via Office.js |
| **Shared Runtime** | A single JavaScript runtime shared between custom functions and task pane |
| **Dynamic Array / Spill** | An Excel feature where a formula result fills multiple cells automatically |
| **TTL** | Time-To-Live — duration a cached entry remains valid |
| **LLM** | Large Language Model |
| **Task Pane** | A sidebar panel in Excel for add-in UI |

### 1.4 References

- Microsoft Office.js Custom Functions: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview
- Anthropic Claude API: https://docs.anthropic.com/en/docs/api-reference
- OpenAI Chat Completions API: https://platform.openai.com/docs/api-reference/chat
- Microsoft COPILOT function documentation (source specification)

---

## 2. Overall Description

### 2.1 Product Perspective

Copilot Clown is a standalone Excel Web Add-in distributed via sideloading or organizational deployment. It operates entirely client-side — API calls are made directly from the user's browser/Excel runtime to the respective AI provider endpoints. No backend server is required.

### 2.2 User Characteristics

Target users are Excel power users who:
- Have API accounts with Anthropic and/or OpenAI
- Want AI-assisted text generation, classification, summarization, and data enrichment within Excel
- Want to control costs and model selection (unlike Microsoft's locked-in COPILOT function)

### 2.3 Constraints

- **Platform:** Excel on Windows, Mac, and Excel Online (Microsoft 365)
- **Runtime:** Office.js Shared Runtime (required for localStorage access in custom functions)
- **Network:** Requires internet connectivity to reach AI provider APIs
- **CORS:** AI provider APIs must support CORS or the add-in must use a proxy (Anthropic API does not allow direct browser requests — a lightweight proxy or the `anthropic-dangerous-direct-browser-access` header is required)
- **Storage:** localStorage is limited to ~5–10 MB per origin depending on browser

### 2.4 Assumptions

- Users have valid API keys for at least one supported provider
- Users understand that AI responses are non-deterministic and may vary
- Excel supports the Custom Functions API (Microsoft 365, Excel 2021+)

---

## 3. Functional Requirements

### 3.1 Custom Function: `=USEAI()`

#### 3.1.1 Syntax

```
=USEAI(prompt_part1, [context1], [prompt_part2], [context2], ...)
```

#### 3.1.2 Arguments

| Argument | Required | Type | Description |
|----------|----------|------|-------------|
| `prompt_part1` | Yes | String | Text describing the task or question |
| `context1` | No | Range or String | Cell reference or range providing data context |
| `prompt_part2` | No | String | Additional prompt text |
| `context2` | No | Range or String | Additional context |
| `...` | No | Alternating String/Range | Additional prompt parts and contexts |

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

- The function is **asynchronous** — Excel shows `#BUSY!` while waiting
- Results are **non-deterministic** — recalculating may produce different output
- The function checks the cache before making an API call
- If no API key is configured, returns a descriptive error string

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
| `claude-opus-4-20250514` | Claude Opus 4 |
| `claude-sonnet-4-20250514` | Claude Sonnet 4 |
| `claude-haiku-3-5-20241022` | Claude Haiku 3.5 |

**OpenAI Models:**

| Model ID | Display Name |
|----------|-------------|
| `gpt-4.1` | GPT-4.1 |
| `gpt-4.1-mini` | GPT-4.1 Mini |
| `gpt-4.1-nano` | GPT-4.1 Nano |
| `gpt-4o` | GPT-4o |
| `gpt-4o-mini` | GPT-4o Mini |
| `o3-mini` | o3-mini |

The model list is hardcoded but designed for easy extension via a registry pattern.

#### 3.2.3 Provider Switching

- The user selects the active provider and model from the task pane
- The selection persists across sessions via localStorage
- All `=USEAI()` calls use the currently active provider/model
- Changing provider/model does **not** automatically recalculate existing formulas

### 3.3 API Key Management

#### 3.3.1 Storage
- API keys are stored in localStorage with Base64 encoding (obfuscation, not encryption)
- Keys are stored per-provider: `aillm_key_anthropic`, `aillm_key_openai`

#### 3.3.2 Task Pane UI
- Text input fields for each provider's API key (masked with `type="password"`)
- "Save" button persists keys to localStorage
- "Test" button makes a minimal API call to validate the key
- Visual indicator (green/red) showing key validity status

### 3.4 Task Pane

The task pane provides a settings panel with three sections:

#### 3.4.1 API Keys Section
- Input fields for Claude and OpenAI API keys
- Save and test functionality per key

#### 3.4.2 Model Selection Section
- Provider toggle (Claude / OpenAI)
- Model dropdown filtered by selected provider
- Shows model capabilities summary (context window, pricing tier)

#### 3.4.3 Cache Management Section
- Display: total cached entries, total cache size, cache hit rate
- "Clear All Cache" button
- TTL configuration input (default: 24 hours)
- Toggle to enable/disable caching globally

### 3.5 Rate Limiting

- Client-side rate limiter: configurable limit (default 100 calls per 10 minutes)
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

- API keys are stored locally only (never transmitted to third parties)
- API calls go directly from the client to the AI provider endpoints
- No telemetry, analytics, or data collection
- Prompts and context data are sent only to the user-selected AI provider
- The add-in has no server-side component

### 4.3 Compatibility

| Platform | Support Level |
|----------|--------------|
| Excel for Windows (Microsoft 365) | Full |
| Excel for Mac (Microsoft 365) | Full |
| Excel Online (web) | Full |
| Excel 2021 (perpetual) | Full (if Custom Functions API supported) |
| Excel 2019 and earlier | Not supported |

### 4.4 Usability

- Function autocomplete in the formula bar with parameter hints
- Descriptive error messages (not just `#VALUE!`)
- Task pane accessible from the ribbon (Home tab or dedicated tab)

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
anthropic-dangerous-direct-browser-access: true
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
| No API key configured | `"Error: No API key configured for <provider>. Open the task pane to add your key."` | — |
| Empty prompt | `"Error: Prompt cannot be empty."` | `#VALUE!` |
| Prompt too large (>100,000 chars) | `"Error: Prompt too large. Reduce prompt or context size."` | `#VALUE!` |
| Rate limit exceeded (client-side) | `"Error: Rate limit exceeded. Wait and try again."` | — |
| API rate limit (429) | `"Error: API rate limit reached. Wait and retry."` | — |
| API authentication error (401/403) | `"Error: Invalid API key for <provider>. Check your key in the task pane."` | — |
| Network timeout | `"Error: Request timed out. Try a simpler prompt or check your connection."` | — |
| Network error | `"Error: Network error. Check your internet connection."` | — |
| API server error (5xx) | `"Error: <Provider> API is unavailable. Try again later."` | — |
| Unknown error | `"Error: Unexpected error occurred."` | — |

### 7.2 Error Strategy

Custom functions return descriptive error strings rather than Excel error codes where possible, because string errors are more informative to the user. Excel error codes (`#VALUE!`, `#CONNECT!`) are used when required by the Office.js custom functions framework.

---

## 8. Caching Specification

### 8.1 Cache Key Generation

```
cacheKey = SHA-256( JSON.stringify({
  provider: "anthropic" | "openai",
  model: "<model_id>",
  prompt: "<full_constructed_prompt>"
}) )
```

The hash is hex-encoded and prefixed: `aillm_cache_<hex_hash>`

### 8.2 Cache Entry Format

```json
{
  "response": "<ai_response_text>",
  "timestamp": 1709337600000,
  "provider": "anthropic",
  "model": "claude-sonnet-4-20250514",
  "promptPreview": "Classify [values] into..."
}
```

### 8.3 Cache Behavior

| Scenario | Action |
|----------|--------|
| Cache hit (within TTL) | Return cached response; no API call |
| Cache hit (expired TTL) | Delete entry; make new API call; cache result |
| Cache miss | Make API call; cache result |
| Cache disabled (user toggle) | Always make API call; do not store |
| localStorage full | Evict oldest entries (LRU) until space is available |

### 8.4 Cache Configuration

| Setting | Default | Range |
|---------|---------|-------|
| TTL | 24 hours | 1 minute – 30 days |
| Caching enabled | `true` | `true` / `false` |
| Max cache entries | 1000 | 100 – 10,000 |

### 8.5 Cache Management (Task Pane)

- **Stats display:** Total entries, total size (KB), hit rate (hits / total lookups)
- **Clear all:** Removes all `aillm_cache_*` keys from localStorage
- **TTL slider:** Adjust cache duration
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

## Appendix B: localStorage Key Map

| Key | Purpose |
|-----|---------|
| `aillm_key_anthropic` | Anthropic API key (Base64 encoded) |
| `aillm_key_openai` | OpenAI API key (Base64 encoded) |
| `aillm_provider` | Active provider (`"anthropic"` or `"openai"`) |
| `aillm_model` | Active model ID |
| `aillm_cache_enabled` | Cache toggle (`"true"` / `"false"`) |
| `aillm_cache_ttl` | Cache TTL in milliseconds |
| `aillm_cache_max` | Max cache entries |
| `aillm_cache_<hash>` | Individual cache entries |
| `aillm_rate_window` | Rate limiter window data |
