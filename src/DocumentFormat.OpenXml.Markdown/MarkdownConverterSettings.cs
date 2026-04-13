using System;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Mode to determine how extracted images are handled during markdown conversion.
/// </summary>
public enum ImageExportMode
{
    /// <summary>
    /// Embed images directly into markdown using a Base64 data URI wrapper.
    /// Example: ![alt](data:image/png;base64,.....)
    /// </summary>
    Base64,

    /// <summary>
    /// Export extracted images out into the configured AssetExportDirectory.
    /// Example outputs standard files and emits ![alt](./assets/image.png)
    /// </summary>
    ExportToFolder,

    /// <summary>
    /// Discard all image data and emit nothing (or a placeholder text).
    /// </summary>
    Ignore,
}

/// <summary>
/// Configuration payload for open xml markdown conversions.
/// </summary>
public class MarkdownConverterSettings
{
    /// <summary>
    /// Gets or sets how extracted images are handled. Default is <see cref="ImageExportMode.Base64"/>.
    /// </summary>
    public ImageExportMode ImageExportMode { get; set; } = ImageExportMode.Base64;

    /// <summary>
    /// Gets or sets for ImageExportMode.ExportToFolder, the target local directory where files will be created.
    /// </summary>
    public string? AssetExportDirectory { get; set; }

    /// <summary>
    /// Gets or sets the base URL path constructed before the file name when using ExportToFolder. Defaults to empty for
    /// relative links.
    /// </summary>
    public string? AssetLinkUrlPrefix { get; set; }

    public void Validate()
    {
        if (this.ImageExportMode == ImageExportMode.ExportToFolder && string.IsNullOrWhiteSpace(this.AssetExportDirectory))
        {
            throw new InvalidOperationException("AssetExportDirectory must be provided when ImageExportMode is set to ExportToFolder.");
        }
    }
}
