# TODO: File & Folder Content Attachment Feature

**Branch:** `feature/file-content-attachment`
**Created:** 2026-03-06

## Phase 1: File Content Extraction & Multimodal (DONE)

- [x] Create feature branch
- [x] Create this TODO tracking file
- [x] Create `Models/ContentModels.cs` (Attachment, ResolvedPrompt)
- [x] Create `Services/ContentCache.cs` (file content caching)
- [x] Create `Services/ContentExtractor.cs` (text extraction + OCR)
- [x] Create `Services/FileResolver.cs` (detection + orchestration)
- [x] Add NuGet packages to `.csproj` (OpenXml, PdfPig, Tesseract)
- [x] Modify `ILlmProvider` + `ClaudeProvider` + `OpenAIProvider` for multimodal
- [x] Modify `UseAiFunction.CallLlm()` to integrate FileResolver
- [x] Update `docs/SRS.md` with file attachment requirements
- [x] Update `CLAUDE.md` with new services
- [x] Build and verify — **0 warnings, 0 errors**

## Phase 2: API File Upload Caching (DONE)

Upload files to provider APIs once, reference by `file_id` in subsequent requests.

- [x] Create `Services/FileUploadCache.cs` (provider + path → file_id)
- [x] Add `RemoteFileId` and `RawBytes` to `Attachment` model
- [x] Add `UploadFileAsync` to `ILlmProvider` interface
- [x] Implement Claude Files API upload (`POST /v1/files`, beta header)
- [x] Implement OpenAI Files API upload (`POST /v1/files`, purpose: `user_data`)
- [x] Update `ClaudeProvider.BuildMultimodalContent()` — use `file_id` when available
- [x] Update `OpenAIProvider.BuildMultimodalContent()` — use `file_id` when available
- [x] Add Claude prompt caching (`cache_control` blocks on content)
- [x] Integrate file upload into `UseAiFunction.CallLlm()` flow
- [x] Add "Clear File Cache" button to `SettingsForm.cs` (AI Settings dialog)
- [x] Update `docs/SRS.md` with API file caching section
- [x] Update `CLAUDE.md`
- [x] Build and verify — **0 warnings, 0 errors**

## Three-Layer Caching Architecture

| Layer | Cache | Key | TTL | Purpose |
|-------|-------|-----|-----|---------|
| 1 | `ContentCache` | path + LastWriteTimeUtc | Infinite (local), 24h (URL) | Avoid re-extracting file content |
| 2 | `FileUploadCache` | provider + path + LastWriteTimeUtc | Infinite (local), 24h (URL) | Avoid re-uploading files to API |
| 3 | `CacheService` | SHA-256(provider\|model\|prompt+attachments) | Configurable (default 24h) | Avoid re-calling API for same prompt |

## API File Upload Details

**Claude (Anthropic):**
- `POST /v1/files` — upload, returns `file_id`
- Beta header: `anthropic-beta: files-api-2025-04-14`
- Reference: `{"type":"document","source":{"type":"file","file_id":"..."}}`
- Images: `{"type":"image","source":{"type":"file","file_id":"..."}}`
- Prompt caching: `{"cache_control":{"type":"ephemeral"}}` on content blocks

**OpenAI:**
- `POST /v1/files` — upload with `purpose: "user_data"`, returns `file_id`
- Reference: `{"type":"file","file":{"file_id":"..."}}`

## Design Decisions

- File paths in prompts enclosed in `{...}` for reliable detection
- File content cache persists until user clears (no auto-expiration for local files)
- Cache key includes `LastWriteTimeUtc` — modified files auto-invalidate
- OCR fallback via Tesseract for scanned PDF pages (< 10 chars extracted text)
- API file upload attempted first; falls back to inline base64/text on upload failure
- Claude prompt caching added for 90% cost reduction on repeated file references
