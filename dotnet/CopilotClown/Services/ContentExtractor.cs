using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using CopilotClown.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Tesseract;
using UglyToad.PdfPig;

namespace CopilotClown.Services;

public class ContentExtractor
{
    private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };

    private static readonly HashSet<string> ImageExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp" };

    private static readonly HashSet<string> DocExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { ".docx", ".pdf" };

    private static readonly HashSet<string> PlainTextExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { ".txt", ".csv", ".md", ".json", ".xml", ".html", ".htm", ".log", ".yaml", ".yml" };

    private const int MaxPlainTextBytes = 1 * 1024 * 1024; // 1 MB
    private const int MaxImageBytes = 20 * 1024 * 1024;    // 20 MB
    private const int MaxDownloadBytes = 50 * 1024 * 1024;  // 50 MB
    private const int MinOcrTextLength = 10; // Minimum chars per page before OCR fallback

    public static bool IsSupported(string extension)
    {
        return ImageExtensions.Contains(extension)
            || DocExtensions.Contains(extension)
            || PlainTextExtensions.Contains(extension);
    }

    public Attachment Extract(string filePath)
    {
        var ext = Path.GetExtension(filePath);
        var fileName = Path.GetFileName(filePath);

        if (ImageExtensions.Contains(ext))
            return ReadImage(filePath, fileName);

        if (ext.Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            // RawBytes not stored here — UploadAttachments reads from disk when needed
            return new Attachment(AttachmentType.Text, ExtractDocx(filePath), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filePath, fileName);
        }

        if (ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            return new Attachment(AttachmentType.Text, ExtractPdf(filePath), "application/pdf", filePath, fileName);

        if (PlainTextExtensions.Contains(ext))
            return new Attachment(AttachmentType.Text, ReadPlainText(filePath), GetMimeType(ext), filePath, fileName);

        return null;
    }

    public Attachment ExtractFromBytes(byte[] data, string fileName, string contentType)
    {
        var ext = Path.GetExtension(fileName);

        // Images: base64 directly from bytes
        if (ImageExtensions.Contains(ext) || contentType?.StartsWith("image/") == true)
        {
            if (data.Length > MaxImageBytes) return null;
            var mime = !string.IsNullOrEmpty(contentType) ? contentType : GetMimeType(ext);
            return new Attachment(AttachmentType.Image, Convert.ToBase64String(data), mime, fileName, fileName)
            {
                RawBytes = data
            };
        }

        // Documents/text: write to temp file, extract, delete
        var tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ext);
        try
        {
            File.WriteAllBytes(tempFile, data);
            var result = Extract(tempFile);
            if (result != null)
            {
                var att = new Attachment(result.Type, result.Content, result.MimeType, fileName, fileName)
                {
                    RawBytes = data
                };
                return att;
            }
            return null;
        }
        finally
        {
            try { File.Delete(tempFile); } catch { }
        }
    }

    public (byte[] data, string contentType, string fileName) DownloadUrl(string url)
    {
        var response = Http.GetAsync(url).GetAwaiter().GetResult();
        response.EnsureSuccessStatusCode();

        var contentType = response.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
        var data = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();

        if (data.Length > MaxDownloadBytes)
            throw new InvalidOperationException($"Downloaded file exceeds {MaxDownloadBytes / 1024 / 1024}MB limit.");

        // Determine filename from URL or Content-Disposition
        var fileName = "download";
        var disposition = response.Content.Headers.ContentDisposition;
        if (disposition?.FileName != null)
        {
            fileName = disposition.FileName.Trim('"');
        }
        else
        {
            var uri = new Uri(url);
            var lastSegment = uri.Segments.LastOrDefault()?.Trim('/');
            if (!string.IsNullOrEmpty(lastSegment) && lastSegment.Contains("."))
                fileName = lastSegment;
            else
                fileName += GetExtensionFromMime(contentType);
        }

        return (data, contentType, fileName);
    }

    private static string ExtractDocx(string filePath)
    {
        var sb = new StringBuilder();
        using (var doc = WordprocessingDocument.Open(filePath, false))
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return "";

            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                var text = paragraph.InnerText;
                if (!string.IsNullOrWhiteSpace(text))
                    sb.AppendLine(text);
            }
        }
        return sb.ToString().Trim();
    }

    private static string ExtractPdf(string filePath)
    {
        var sb = new StringBuilder();
        var ocrPages = new List<int>();

        using (var doc = PdfDocument.Open(filePath))
        {
            for (int i = 1; i <= doc.NumberOfPages; i++)
            {
                var page = doc.GetPage(i);
                var pageText = page.Text;

                if (string.IsNullOrWhiteSpace(pageText) || pageText.Trim().Length < MinOcrTextLength)
                {
                    ocrPages.Add(i);
                    sb.AppendLine($"[Page {i}: scanned — see OCR below]");
                }
                else
                {
                    sb.AppendLine($"--- Page {i} ---");
                    sb.AppendLine(pageText.Trim());
                }
                sb.AppendLine();
            }
        }

        // OCR fallback for scanned pages
        if (ocrPages.Count > 0)
        {
            var ocrText = OcrPdfPages(filePath, ocrPages);
            if (!string.IsNullOrEmpty(ocrText))
                sb.AppendLine(ocrText);
        }

        return sb.ToString().Trim();
    }

    private static string OcrPdfPages(string pdfPath, List<int> pageNumbers)
    {
        var sb = new StringBuilder();

        try
        {
            // Tesseract data path: look next to the assembly, then in a well-known location
            var tessDataPath = FindTessDataPath();
            if (tessDataPath == null)
                return "[OCR unavailable: Tesseract language data not found]";

            using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
            using (var pdfDoc = PdfDocument.Open(pdfPath))
            {
                foreach (var pageNum in pageNumbers)
                {
                    var page = pdfDoc.GetPage(pageNum);

                    // Get page images embedded in the PDF
                    var images = page.GetImages().ToList();
                    if (images.Count == 0) continue;

                    sb.AppendLine($"--- Page {pageNum} (OCR) ---");

                    foreach (var pdfImage in images)
                    {
                        try
                        {
                            var imageBytes = pdfImage.RawBytes.ToArray();
                            using (var pix = Pix.LoadFromMemory(imageBytes))
                            using (var result = engine.Process(pix))
                            {
                                var text = result.GetText()?.Trim();
                                if (!string.IsNullOrEmpty(text))
                                    sb.AppendLine(text);
                            }
                        }
                        catch
                        {
                            // Skip unprocessable images silently
                        }
                    }
                    sb.AppendLine();
                }
            }
        }
        catch (Exception ex)
        {
            return $"[OCR error: {ex.Message}]";
        }

        return sb.ToString().Trim();
    }

    private static string FindTessDataPath()
    {
        // 1. Next to the executing assembly
        var assemblyDir = Path.GetDirectoryName(typeof(ContentExtractor).Assembly.Location);
        if (assemblyDir != null)
        {
            var local = Path.Combine(assemblyDir, "tessdata");
            if (Directory.Exists(local)) return local;
        }

        // 2. TESSDATA_PREFIX environment variable
        var envPath = Environment.GetEnvironmentVariable("TESSDATA_PREFIX");
        if (!string.IsNullOrEmpty(envPath) && Directory.Exists(envPath))
            return envPath;

        // 3. Common Windows install location
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var commonPath = Path.Combine(programFiles, "Tesseract-OCR", "tessdata");
        if (Directory.Exists(commonPath)) return commonPath;

        return null;
    }

    public static int GetPdfPageCount(byte[] fileBytes)
    {
        using (var doc = PdfDocument.Open(fileBytes))
            return doc.NumberOfPages;
    }

    private static string ReadPlainText(string filePath)
    {
        var info = new FileInfo(filePath);
        if (info.Length > MaxPlainTextBytes)
            return File.ReadAllText(filePath, Encoding.UTF8).Substring(0, MaxPlainTextBytes);
        return File.ReadAllText(filePath, Encoding.UTF8);
    }

    private static Attachment ReadImage(string filePath, string fileName)
    {
        var info = new FileInfo(filePath);
        if (info.Length > MaxImageBytes) return null;

        var bytes = File.ReadAllBytes(filePath);
        var ext = Path.GetExtension(filePath);
        return new Attachment(AttachmentType.Image, Convert.ToBase64String(bytes), GetMimeType(ext), filePath, fileName);
    }

    private static string GetMimeType(string extension)
    {
        switch (extension.ToLowerInvariant())
        {
            case ".png": return "image/png";
            case ".jpg": case ".jpeg": return "image/jpeg";
            case ".gif": return "image/gif";
            case ".webp": return "image/webp";
            case ".bmp": return "image/bmp";
            case ".pdf": return "application/pdf";
            case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            case ".txt": return "text/plain";
            case ".csv": return "text/csv";
            case ".md": return "text/markdown";
            case ".json": return "application/json";
            case ".xml": return "application/xml";
            case ".html": case ".htm": return "text/html";
            case ".log": return "text/plain";
            case ".yaml": case ".yml": return "application/yaml";
            default: return "application/octet-stream";
        }
    }

    private static string GetExtensionFromMime(string mimeType)
    {
        switch (mimeType?.ToLowerInvariant())
        {
            case "image/png": return ".png";
            case "image/jpeg": return ".jpg";
            case "image/gif": return ".gif";
            case "image/webp": return ".webp";
            case "image/bmp": return ".bmp";
            case "application/pdf": return ".pdf";
            case "application/vnd.openxmlformats-officedocument.wordprocessingml.document": return ".docx";
            case "text/plain": return ".txt";
            case "text/csv": return ".csv";
            case "text/html": return ".html";
            case "application/json": return ".json";
            case "application/xml": return ".xml";
            default: return ".bin";
        }
    }
}
