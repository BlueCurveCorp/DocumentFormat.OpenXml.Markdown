using System;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Internal parser to convert PresentationDocument to Markdown.
/// </summary>
internal static class PresentationParser
{
    public static string Parse(PresentationDocument document, MarkdownConverterSettings settings)
    {
        var sb = new StringBuilder();
        var presentationPart = document.PresentationPart;

        if (presentationPart?.Presentation?.SlideIdList is null)
        {
            return string.Empty;
        }

        var slideIndex = 1;

        foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
        {
            if (slideId.RelationshipId is null)
            {
                continue;
            }

            var slidePart = presentationPart.GetPartById(slideId.RelationshipId.Value!) as SlidePart;
            if (slidePart?.Slide is null)
            {
                continue;
            }

            // Try to find the slide title
            string? slideTitle = null;
            var titleShape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
                .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.Title ||
                                    s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.CenteredTitle);

            if (titleShape is not null && titleShape.TextBody is not null)
            {
                slideTitle = titleShape.TextBody.InnerText;
            }

            sb.AppendLine($"## {(string.IsNullOrEmpty(slideTitle) ? $"Slide {slideIndex}" : slideTitle)}");
            sb.AppendLine();

            var shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>();

            foreach (var shape in shapes)
            {
                // Skip title shape as we already used it
                if (shape == titleShape)
                {
                    continue;
                }

                var textBody = shape.TextBody;
                if (textBody is not null)
                {
                    foreach (var paragraph in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
                    {
                        ParseDrawingParagraph(paragraph, sb);
                    }
                }
            }

            // Extract tables from GraphicFrames
            var graphicFrames = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>();
            foreach (var frame in graphicFrames)
            {
                var table = frame.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();
                if (table is not null)
                {
                    ParseDrawingTable(table, sb);
                }
            }


            // Extract any pics if needed
            var pics = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().ToList();
            if (pics.Count > 0 && settings.ImageExportMode != ImageExportMode.Ignore)
            {
                sb.AppendLine();
                foreach (var pic in pics)
                {
                    var blip = pic.BlipFill?.Blip;
                    if (blip?.Embed?.Value is not null)
                    {
                        if (slidePart.GetPartById(blip.Embed.Value) is ImagePart imagePart)
                        {
                            sb.AppendLine(ImageExtractor.ExtractImage(imagePart, settings, $"slide{slideIndex}_img"));
                            sb.AppendLine();
                        }
                    }
                }
            }

            sb.AppendLine();
            slideIndex++;
        }

        return sb.ToString().TrimEnd();
    }

    private static void ParseDrawingParagraph(DocumentFormat.OpenXml.Drawing.Paragraph paragraph, StringBuilder sb)
    {
        var isBullet = paragraph.ParagraphProperties?.GetFirstChild<BulletFont>() is not null ||
                       paragraph.ParagraphProperties?.GetFirstChild<CharacterBullet>() is not null;

        if (isBullet)
        {
            var level = paragraph.ParagraphProperties?.Level?.Value ?? 0;
            for (var i = 0; i < level; i++)
            {
                sb.Append("  ");
            }

            sb.Append("- ");
        }

        foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>())
        {
            var isBold = run.RunProperties?.Bold?.Value ?? false;
            var isItalic = run.RunProperties?.Italic?.Value ?? false;

            if (isBold)
            {
                sb.Append("**");
            }

            if (isItalic)
            {
                sb.Append('*');
            }

            sb.Append(run.Text?.Text);

            if (isItalic)
            {
                sb.Append('*');
            }

            if (isBold)
            {
                sb.Append("**");
            }
        }

        sb.AppendLine();
        sb.AppendLine();
    }

    private static void ParseDrawingTable(DocumentFormat.OpenXml.Drawing.Table table, StringBuilder sb)
    {
        var rows = table.Elements<DocumentFormat.OpenXml.Drawing.TableRow>().ToList();
        if (rows.Count == 0)
        {
            return;
        }

        // Header
        var firstRow = rows[0];
        sb.Append('|');
        foreach (var cell in firstRow.Elements<DocumentFormat.OpenXml.Drawing.TableCell>())
        {
            var text = cell.TextBody?.InnerText.Trim() ?? string.Empty;
            sb.Append(' ');
            sb.Append(text.Replace("|", "\\|", StringComparison.Ordinal));
            sb.Append(" |");
        }

        sb.AppendLine();

        // Separator
        sb.Append('|');

        foreach (var _ in firstRow.Elements<DocumentFormat.OpenXml.Drawing.TableCell>())
        {
            sb.Append(" --- |");
        }

        sb.AppendLine();

        // Data rows
        for (var i = 1; i < rows.Count; i++)
        {
            sb.Append('|');

            foreach (var cell in rows[i].Elements<DocumentFormat.OpenXml.Drawing.TableCell>())
            {
                var text = cell.TextBody?.InnerText.Trim() ?? string.Empty;
                sb.Append(' ');
                sb.Append(text.Replace("|", "\\|", StringComparison.Ordinal));
                sb.Append(" |");
            }

            sb.AppendLine();
        }
        sb.AppendLine();
    }
}
