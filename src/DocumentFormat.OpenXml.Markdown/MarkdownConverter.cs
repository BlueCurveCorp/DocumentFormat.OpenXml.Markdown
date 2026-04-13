using System;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Entry point for converting OpenXml documents to Markdown.
/// </summary>
public static class MarkdownConverter
{
    /// <summary>
    /// Converts the contents of an OpenXML document stream to a Markdown-formatted string.
    /// </summary>
    /// <remarks>
    /// The method automatically detects the type of OpenXML document and applies the appropriate conversion. The stream
    /// must support reading and be positioned at the start of the document.
    /// </remarks>
    /// <param name="package">The OpenXmlPackage representing the document to convert.</param>
    /// <param name="settings">
    /// Optional settings that control the Markdown conversion process. If null, default settings are used.
    /// </param>
    /// <returns>A string containing the Markdown representation of the document's contents.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown if the provided package is null.
    /// </exception>
    /// <exception cref="NotSupportedException">
    /// Thrown if the provided package does not contain a recognized OpenXML document type (Word, Excel, or PowerPoint).
    /// </exception>
    public static string ConvertToMarkdown(OpenXmlPackage package, MarkdownConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(package);

        var effectiveSettings = settings ?? new MarkdownConverterSettings();

        return package switch
        {
            WordprocessingDocument wordDoc => DocumentParser.Parse(wordDoc, effectiveSettings),
            PresentationDocument pptDoc => PresentationParser.Parse(pptDoc, effectiveSettings),
            SpreadsheetDocument excelDoc => SpreadsheetParser.Parse(excelDoc, effectiveSettings),
            _ => throw new NotSupportedException("Unsupported document type: " + package.GetType().Name),
        };
    }

    /// <summary>
    /// Asynchronously converts the contents of an OpenXML document stream to a Markdown-formatted string.
    /// </summary>
    /// <remarks>
    /// The method automatically detects the type of OpenXML document and applies the appropriate conversion. The stream
    /// must support reading and be positioned at the start of the document.
    /// </remarks>
    /// <param name="package">The OpenXmlPackage representing the document to convert.</param>
    /// <param name="settings">
    /// Optional settings that control the Markdown conversion process. If null, default settings are used.
    /// </param>
    /// <returns>A string containing the Markdown representation of the document's contents.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown if the provided package is null.
    /// </exception>
    /// <exception cref="NotSupportedException">
    /// Thrown if the provided package does not contain a recognized OpenXML document type (Word, Excel, or PowerPoint).
    /// </exception>
    public async static Task<string> ConvertToMarkdownAsync(OpenXmlPackage package, MarkdownConverterSettings? settings = null)
    {
        return await Task.Run(() => ConvertToMarkdown(package, settings)).ConfigureAwait(false);
    }
}
