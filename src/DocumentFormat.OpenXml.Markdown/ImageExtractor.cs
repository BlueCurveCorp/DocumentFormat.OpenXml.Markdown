using System;
using System.IO;

using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Helper to extract images from OpenXml parts.
/// </summary>
internal static class ImageExtractor
{
    public static string ExtractImage(ImagePart imagePart, MarkdownConverterSettings settings, string baseName)
    {
        if (settings.ImageExportMode == ImageExportMode.Ignore)
        {
            return string.Empty;
        }

        var extension = GetExtension(imagePart.ContentType);
        var fileName = string.Concat(baseName, "_", Guid.NewGuid().ToString("N").AsSpan(0, 8), extension);

        if (settings.ImageExportMode == ImageExportMode.ExportToFolder && !string.IsNullOrEmpty(settings.AssetExportDirectory))
        {
            if (!Directory.Exists(settings.AssetExportDirectory))
            {
                Directory.CreateDirectory(settings.AssetExportDirectory!);
            }

            var filePath = Path.Combine(settings.AssetExportDirectory!, fileName);
            using (var stream = imagePart.GetStream())
            using (var fileStream = File.Create(filePath))
            {
                stream.CopyTo(fileStream);
            }

            var url = string.Concat(settings.AssetLinkUrlPrefix ?? string.Empty, fileName);
            return "![" + baseName + "](" + url + ")";
        }

        return "![" + baseName + "](embedded_image)";
    }

    private static string GetExtension(string contentType)
    {
        return contentType switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/x-png" => ".png",
            "image/tiff" => ".tiff",
            _ => ".bin",
        };
    }
}
