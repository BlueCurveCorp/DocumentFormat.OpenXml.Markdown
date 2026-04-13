using System;
using System.IO;
using System.IO.Packaging;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Provides methods to convert OpenXML documents into Markdown format.
/// </summary>
public static class MarkdownConverter
{
    private static readonly MarkdownConverterSettings DefaultSettings = new MarkdownConverterSettings();

    /// <summary>
    /// Converts the contents of a WordprocessingDocument to a Markdown-formatted string.
    /// </summary>
    /// <remarks>The conversion process uses the specified settings to control formatting and output. If no
    /// settings are provided, default conversion options are applied.</remarks>
    /// <param name="document">The WordprocessingDocument to convert. Cannot be null.</param>
    /// <param name="settings">Optional settings that control the Markdown conversion process. If null, default settings are used.</param>
    /// <returns>A string containing the Markdown representation of the document's contents.</returns>
    public static string ConvertToMarkdown(WordprocessingDocument document, MarkdownConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(document);

        var s = settings ?? DefaultSettings;
        s.Validate();
        return DocumentParser.Parse(document, s);
    }

    /// <summary>
    /// Converts the specified spreadsheet document to a Markdown-formatted string.
    /// </summary>
    /// <remarks>Use this method to generate Markdown output from an Open XML spreadsheet. The conversion can
    /// be customized by providing specific settings; otherwise, default conversion behavior is applied.</remarks>
    /// <param name="document">The spreadsheet document to convert. Cannot be null.</param>
    /// <param name="settings">Optional settings that control the Markdown conversion process. If null, default settings are used.</param>
    /// <returns>A string containing the Markdown representation of the spreadsheet document.</returns>
    public static string ConvertToMarkdown(SpreadsheetDocument document, MarkdownConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(document);

        var s = settings ?? DefaultSettings;
        s.Validate();
        return SpreadsheetParser.Parse(document, s);
    }

    /// <summary>
    /// Converts the specified PowerPoint presentation to a Markdown-formatted string.
    /// </summary>
    /// <remarks>The conversion process uses the provided settings to control formatting and content
    /// extraction. If no settings are specified, default conversion options are applied.</remarks>
    /// <param name="document">The PowerPoint presentation to convert. Cannot be null.</param>
    /// <param name="settings">Optional settings that control the Markdown conversion process. If null, default settings are used.</param>
    /// <returns>A string containing the Markdown representation of the presentation.</returns>
    public static string ConvertToMarkdown(PresentationDocument document, MarkdownConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(document);

        var s = settings ?? DefaultSettings;
        s.Validate();
        return PresentationParser.Parse(document, s);
    }

    /// <summary>
    /// Converts the contents of an OpenXML document stream to a Markdown-formatted string.
    /// </summary>
    /// <remarks>The method automatically detects the type of OpenXML document and applies the appropriate
    /// conversion. The stream must support reading and be positioned at the start of the document.</remarks>
    /// <param name="stream">The input stream containing the OpenXML document to convert. Must be positioned at the beginning of a valid
    /// Word, Excel, or PowerPoint file. Cannot be null.</param>
    /// <param name="settings">Optional settings that control the Markdown conversion process. If null, default settings are used.</param>
    /// <returns>A string containing the Markdown representation of the document's contents.</returns>
    /// <exception cref="ArgumentException">Thrown if the provided stream does not contain a recognized OpenXML document type (Word, Excel, or PowerPoint).</exception>
    public static string ConvertToMarkdown(Stream stream, MarkdownConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(stream);

        using var package = Package.Open(stream, FileMode.Open, FileAccess.Read);

        // Attempt to open as Word
        if (package.PartExists(new Uri("/word/document.xml", UriKind.Relative)))
        {
            using var doc = WordprocessingDocument.Open(package);
            return ConvertToMarkdown(doc, settings);
        }

        if (package.PartExists(new Uri("/xl/workbook.xml", UriKind.Relative)))
        {
            using var doc = SpreadsheetDocument.Open(package);
            return ConvertToMarkdown(doc, settings);
        }

        if (package.PartExists(new Uri("/ppt/presentation.xml", UriKind.Relative)))
        {
            using var doc = PresentationDocument.Open(package);
            return ConvertToMarkdown(doc, settings);
        }

        throw new ArgumentException("Provided stream is not a recognized OpenXML document type.");
    }

    /// <summary>
    /// Asynchronously converts the contents of a WordprocessingDocument to a Markdown-formatted string.
    /// </summary>
    /// <remarks>The caller is responsible for disposing the WordprocessingDocument after the conversion
    /// completes. The method runs the conversion on a background thread. Thread safety of the document is not
    /// guaranteed if accessed concurrently.</remarks>
    /// <param name="document">The WordprocessingDocument to convert. Cannot be null. The document must be open and readable.</param>
    /// <param name="settings">Optional settings that control the Markdown conversion process. If null, default settings are used.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains a string with the Markdown
    /// representation of the document.</returns>
    public static Task<string> ConvertToMarkdownAsync(WordprocessingDocument document, MarkdownConverterSettings? settings = null)
        => Task.Run(() => ConvertToMarkdown(document, settings));

    /// <summary>
    /// Asynchronously converts the contents of a SpreadsheetDocument to a Markdown-formatted string.
    /// </summary>
    /// <remarks>
    /// The caller is responsible for disposing the SpreadsheetDocument after the conversion completes. The method runs
    /// the conversion on a background thread. Thread safety of the document is not guaranteed if accessed concurrently.
    /// </remarks>
    /// <param name="document">
    /// The SpreadsheetDocument to convert. Cannot be null. The document must be open and readable.
    /// </param>
    /// <param name="settings">
    /// Optional settings that control the Markdown conversion process. If null, default settings are used.
    /// </param>
    /// <returns>
    /// A task that represents the asynchronous operation. The task result contains a string with the Markdown
    /// representation of the document.
    /// </returns>
    public static Task<string> ConvertToMarkdownAsync(SpreadsheetDocument document, MarkdownConverterSettings? settings = null)
        => Task.Run(() => ConvertToMarkdown(document, settings));

    /// <summary>
    /// Asynchronously converts the contents of a PresentationDocument to a Markdown-formatted string.
    /// </summary>
    /// <remarks>
    /// The caller is responsible for disposing the PresentationDocument after the conversion completes. The method runs
    /// the conversion on a background thread. Thread safety of the document is not guaranteed if accessed concurrently.
    /// </remarks>
    /// <param name="document">
    /// The PresentationDocument to convert. Cannot be null. The document must be open and readable.
    /// </param>
    /// <param name="settings">
    /// Optional settings that control the Markdown conversion process. If null, default settings are used.
    /// </param>
    /// <returns>
    /// A task that represents the asynchronous operation. The task result contains a string with the Markdown
    /// representation of the document.
    /// </returns>
    public static Task<string> ConvertToMarkdownAsync(PresentationDocument document, MarkdownConverterSettings? settings = null)
        => Task.Run(() => ConvertToMarkdown(document, settings));

    /// <summary>
    /// Asynchronously converts the contents of an OpenXML document stream to a Markdown-formatted string..
    /// </summary>
    /// <remarks>
    /// The caller is responsible for disposing the PresentationDocument after the conversion completes. The method runs
    /// the conversion on a background thread. Thread safety of the document is not guaranteed if accessed concurrently.
    /// </remarks>
    /// <param name="stream">The stream containing the OpenXML document. Cannot be null. The stream must be readable.</param>
    /// <param name="settings">
    /// Optional settings that control the Markdown conversion process. If null, default settings are used.
    /// </param>
    /// <returns>
    /// A task that represents the asynchronous operation. The task result contains a string with the Markdown
    /// representation of the document.
    /// </returns>
    public static Task<string> ConvertToMarkdownAsync(Stream stream, MarkdownConverterSettings? settings = null)
        => Task.Run(() => ConvertToMarkdown(stream, settings));
}
