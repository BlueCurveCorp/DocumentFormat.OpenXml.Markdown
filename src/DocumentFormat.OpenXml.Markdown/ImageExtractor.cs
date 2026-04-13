using System;
using System.IO;

using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Markdown;

internal static class ImageExtractor
{
    public static string ExtractImage(ImagePart imagePart, MarkdownConverterSettings settings, string fallbackName)
    {
        if (settings.ImageExportMode == ImageExportMode.Ignore || imagePart is null)
        {
            return string.Empty;
        }

#pragma warning disable CA1031 // Do not catch general exception types
        try
        {
            using var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            if (settings.ImageExportMode == ImageExportMode.Base64)
            {
                using var memoryStream = new MemoryStream();
                stream.CopyTo(memoryStream);
                var base64String = Convert.ToBase64String(memoryStream.ToArray());

                // Determine mime type from content type (e.g., image/jpeg)
                var mimeType = imagePart.ContentType ?? "image/png";
                return $"![Image](data:{mimeType};base64,{base64String})";
            }
            else if (settings.ImageExportMode == ImageExportMode.ExportToFolder)
            {
                if (string.IsNullOrWhiteSpace(settings.AssetExportDirectory))
                {
                    return "![Image]"; // Fallback if no directory configured
                }

                // Determine extension and normalized mime type
                var ext = ".bin";
                var mimeType = imagePart.ContentType ?? "application/octet-stream";

                if (mimeType == "image/jpeg") ext = ".jpg";
                else if (mimeType == "image/png") ext = ".png";
                else if (mimeType == "image/gif") ext = ".gif";
                else if (mimeType == "image/svg+xml") ext = ".svg";
                else if (mimeType == "image/webp") ext = ".webp";
                else if (mimeType == "image/tiff") ext = ".tiff";
                else if (mimeType == "image/bmp") ext = ".bmp";
                else if (mimeType == "image/x-icon") ext = ".ico";

                var fileName = $"{fallbackName}_{Guid.NewGuid().ToString("n").Substring(0, 8)}{ext}";

                var exportPath = Path.Combine(settings.AssetExportDirectory, fileName);

                // Ensure directory exists
                Directory.CreateDirectory(settings.AssetExportDirectory);

                using (var fileStream = new FileStream(exportPath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }

                var linkPrefix = settings.AssetLinkUrlPrefix ?? string.Empty;

                if (!string.IsNullOrEmpty(linkPrefix) && !linkPrefix.EndsWith('/'))
                {
                    linkPrefix += "/";
                }

                return $"![Image]({linkPrefix}{fileName})";
            }
        }
        catch
        {
            // If image extraction fails, gracefully fall back
        }
#pragma warning restore CA1031 // Do not catch general exception types

        return "![Image]";
    }
}
