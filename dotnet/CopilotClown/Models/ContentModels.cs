using System.Collections.Generic;

namespace CopilotClown.Models;

public enum AttachmentType
{
    Text,
    Image
}

public class Attachment
{
    public AttachmentType Type { get; }
    public string Content { get; set; } // Text: extracted text. Image: base64-encoded data.
    public string MimeType { get; }     // e.g. "image/png", "application/pdf"
    public string SourcePath { get; }   // Original file path or URL
    public string FileName { get; }     // Just the filename
    public string RemoteFileId { get; set; } // API-uploaded file ID (set after upload)
    public byte[] RawBytes { get; set; }     // Original file bytes (for upload to API)

    public Attachment(AttachmentType type, string content, string mimeType, string sourcePath, string fileName)
    {
        Type = type;
        Content = content;
        MimeType = mimeType;
        SourcePath = sourcePath;
        FileName = fileName;
    }
}

public class ResolvedPrompt
{
    public string CleanText { get; }
    public List<Attachment> Attachments { get; }
    public bool HasAttachments => Attachments.Count > 0;

    public ResolvedPrompt(string cleanText, List<Attachment> attachments)
    {
        CleanText = cleanText;
        Attachments = attachments;
    }

    // Backward compat: plain string prompt with no attachments
    public ResolvedPrompt(string plainPrompt) : this(plainPrompt, new List<Attachment>()) { }
}
