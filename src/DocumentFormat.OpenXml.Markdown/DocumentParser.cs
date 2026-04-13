using System;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Internal parser to convert WordprocessingDocument to Markdown.
/// </summary>
internal static class DocumentParser
{
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static string Parse(WordprocessingDocument document, MarkdownConverterSettings settings)
    {
        var sb = new StringBuilder();
        var mainPart = document.MainDocumentPart;
        var body = mainPart?.Document?.Body;

        if (body is null)
        {
            return string.Empty;
        }

        ParseElement(body, mainPart, settings, sb);

        return sb.ToString().TrimEnd();
    }

    private static void ParseElement(OpenXmlCompositeElement parent, MainDocumentPart? mainPart, MarkdownConverterSettings settings, StringBuilder sb)
    {
        foreach (var element in parent.ChildElements)
        {
            if (element is Paragraph paragraph)
            {
                ParseParagraph(paragraph, mainPart, settings, sb);
            }
            else if (element is Table table)
            {
                ParseTable(table, mainPart, settings, sb);
            }
            else if (element.NamespaceUri == MathNamespace)
            {
                if (element.LocalName == "oMathPara")
                {
                    sb.Append(MathParser.ParseOfficeMathPara(element));
                }
                else if (element.LocalName == "oMath")
                {
                    sb.Append(MathParser.ParseOfficeMath(element));
                }
            }
            else if (element is OpenXmlCompositeElement composite)
            {
                ParseElement(composite, mainPart, settings, sb);
            }
        }
    }

    private static void ParseTable(Table table, MainDocumentPart? mainPart, MarkdownConverterSettings settings, StringBuilder sb)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0)
        {
            return;
        }

        var isFirstRow = true;
        foreach (var row in rows)
        {
            sb.Append('|');
            var cells = row.Elements<TableCell>().ToList();
            foreach (var cell in cells)
            {
                var cellContent = new StringBuilder();
                ParseElement(cell, mainPart, settings, cellContent);

                var text = cellContent.ToString().Trim().Replace("\r", "", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal).Replace("|", "\\|", StringComparison.Ordinal);

                sb.Append($" {text} |");
            }
            sb.AppendLine();

            if (isFirstRow)
            {
                sb.Append('|');
                foreach (var unused in cells)
                {
                    sb.Append(" --- |");
                }
                sb.AppendLine();
                isFirstRow = false;
            }
        }
        sb.AppendLine();
    }

    private static void ParseParagraph(Paragraph paragraph, MainDocumentPart? mainPart, MarkdownConverterSettings settings, StringBuilder sb)
    {
        var paragraphProperties = paragraph.ParagraphProperties;

        var isHeading = false;
        var headingLevel = 0;
        var isList = false;

        if (paragraphProperties is not null)
        {
            var pStyle = paragraphProperties.ParagraphStyleId?.Val?.Value;
            if (pStyle is not null && pStyle.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(pStyle.AsSpan(7), out var level))
                {
                    isHeading = true;
                    headingLevel = level;
                }
            }

            // Detect numbering
            if (paragraphProperties.NumberingProperties is not null)
            {
                isList = true;
            }
        }

        if (isHeading)
        {
            sb.Append(new string('#', headingLevel)).Append(' ');
        }
        else if (isList)
        {
            var listLevel = 0;
            if (paragraphProperties?.NumberingProperties?.NumberingLevelReference?.Val is not null)
            {
                listLevel = paragraphProperties.NumberingProperties.NumberingLevelReference.Val.Value;
            }

            for (var i = 0; i < listLevel; i++)
            {
                sb.Append("  ");
            }

            sb.Append("- ");
        }

        ParseInlineElement(paragraph, mainPart, settings, sb);

        sb.AppendLine();

        // Output double line break for regular paragraphs, single for list items
        if (!isList)
        {
            sb.AppendLine();
        }
    }

    private static void ParseInlineElement(OpenXmlCompositeElement parent, MainDocumentPart? mainPart, MarkdownConverterSettings settings, StringBuilder sb)
    {
        foreach (var child in parent.ChildElements)
        {
            if (child is Run run)
            {
                ParseRun(run, mainPart, settings, sb);
            }
            else if (child is Hyperlink hyperlink)
            {
                ParseHyperlink(hyperlink, mainPart, settings, sb);
            }
            else if (child.NamespaceUri == MathNamespace && child.LocalName == "oMath")
            {
                sb.Append(MathParser.ParseOfficeMath(child));
            }
            else if (child is OpenXmlCompositeElement composite)
            {
                ParseInlineElement(composite, mainPart, settings, sb);
            }
        }
    }

    private static void ParseRun(Run run, MainDocumentPart? mainPart, MarkdownConverterSettings settings, StringBuilder sb)
    {
        var isBold = false;
        var isItalic = false;
        var isStrike = false;

        var rPr = run.RunProperties;
        if (rPr is not null)
        {
            if (rPr.Bold is not null && (rPr.Bold.Val is null || rPr.Bold.Val.Value))
            {
                isBold = true;
            }

            if (rPr.Italic is not null && (rPr.Italic.Val is null || rPr.Italic.Val.Value))
            {
                isItalic = true;
            }

            if (rPr.Strike is not null && (rPr.Strike.Val is null || rPr.Strike.Val.Value))
            {
                isStrike = true;
            }
        }

        // For simplicity, just append text with formats. Real nested format handling can be complex.
        var text = string.Concat(run.Elements<Text>().Select(t => t.Text));

        if (string.IsNullOrEmpty(text) && run.Elements<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any())
        {
            if (settings.ImageExportMode != ImageExportMode.Ignore && mainPart is not null)
            {
                var blips = run.Descendants<DocumentFormat.OpenXml.Drawing.Blip>();
                foreach (var blip in blips)
                {
                    if (blip.Embed?.Value is not null)
                    {
                        if (mainPart.GetPartById(blip.Embed.Value) is ImagePart imagePart)
                        {
                            sb.Append(ImageExtractor.ExtractImage(imagePart, settings, "Image"));
                        }
                    }
                }
            }
            return;
        }

        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        if (isBold)
        {
            sb.Append("**");
        }

        if (isItalic)
        {
            sb.Append('*');
        }

        if (isStrike)
        {
            sb.Append("~~");
        }

        sb.Append(text);

        if (isStrike)
        {
            sb.Append("~~");
        }

        if (isItalic)
        {
            sb.Append('*');
        }

        if (isBold)
        {
            sb.Append("**");
        }
    }

    private static void ParseHyperlink(Hyperlink hyperlink, MainDocumentPart? mainPart, MarkdownConverterSettings settings, StringBuilder sb)
    {
        var url = string.Empty;

        if (hyperlink.Id is not null && mainPart is not null)
        {
            var rel = mainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id);
            if (rel is not null)
            {
                url = rel.Uri.ToString();
            }
        }

        sb.Append('[');

        foreach (var child in hyperlink.Elements<Run>())
        {
            ParseRun(child, mainPart, settings, sb);
        }

        sb.Append("](").Append(url).Append(')');
    }
}
