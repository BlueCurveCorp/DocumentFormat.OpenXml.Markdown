using System;
using System.Globalization;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using D = DocumentFormat.OpenXml.Drawing;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Internal parser to convert PresentationDocument to Markdown.
/// </summary>
internal static class PresentationParser
{
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static string Parse(PresentationDocument document, MarkdownConverterSettings settings)
    {
        var sb = new StringBuilder();
        var presentationPart = document.PresentationPart;

        if (presentationPart is null)
        {
            return string.Empty;
        }

        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<SlideId>();

        if (slideIds is null)
        {
            return string.Empty;
        }

        var slideIndex = 1;
        foreach (var slideId in slideIds)
        {
            if (slideId.RelationshipId is null)
            {
                continue;
            }

            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            var slideTitle = string.Empty;

            var titleShape = slidePart.Slide?.Descendants<Shape>()
                .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.Title ||
                                    s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.CenteredTitle);

            if (titleShape is not null && titleShape.TextBody is not null)
            {
                slideTitle = titleShape.TextBody.InnerText;
            }

            sb.AppendFormat(CultureInfo.InvariantCulture, "## {0}", string.IsNullOrEmpty(slideTitle) ? "Slide " + slideIndex : slideTitle).AppendLine();
            sb.AppendLine();

            var shapes = slidePart!.Slide!.Descendants<Shape>();

            foreach (var shape in shapes)
            {
                if (shape.TextBody is not null)
                {
                    foreach (var paragraph in shape.TextBody.Elements<D.Paragraph>())
                    {
                        ParseDrawingParagraph(paragraph, slidePart, settings, sb);
                    }
                }
            }

            var tables = slidePart.Slide.Descendants<D.Table>();
            foreach (var table in tables)
            {
                ParseDrawingTable(table, sb);
            }

            var pics = slidePart.Slide.Descendants<Picture>().ToList();
            if (pics.Count > 0 && settings.ImageExportMode != ImageExportMode.Ignore)
            {
                foreach (var pic in pics)
                {
                    var blip = pic.BlipFill?.Blip;
                    if (blip?.Embed is not null)
                    {
                        if (slidePart.TryGetPartById(blip.Embed!, out var part) && part is ImagePart imagePart)
                        {
                            sb.AppendLine(ImageExtractor.ExtractImage(imagePart, settings, "ppt_slide_" + slideIndex));
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

#pragma warning disable IDE0060 // Remove unused parameter
    private static void ParseDrawingParagraph(D.Paragraph paragraph, SlidePart slidePart, MarkdownConverterSettings settings, StringBuilder sb)
#pragma warning restore IDE0060 // Remove unused parameter
    {
        var pPr = paragraph.ParagraphProperties;
        var level = pPr?.Level?.Value ?? 0;

        var isBullet = false;

        if (pPr is not null)
        {
            foreach (var child in pPr.ChildElements)
            {
                if (child.LocalName.StartsWith("bu", StringComparison.Ordinal) && child.LocalName != "buNone")
                {
                    isBullet = true;
                    break;
                }
            }
        }

        if (level > 0 || isBullet)
        {
            sb.Append(new string(' ', level * 2)).Append("- ");
        }

        foreach (var child in paragraph.ChildElements)
        {
            if (child is D.Run run)
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
            else if (child.NamespaceUri == MathNamespace && child.LocalName == "oMath")
            {
                sb.Append(MathParser.ParseOfficeMath(child));
            }
        }

        sb.AppendLine();
    }

    private static void ParseDrawingTable(D.Table table, StringBuilder sb)
    {
        var rows = table.Elements<D.TableRow>().ToList();
        if (rows.Count == 0)
        {
            return;
        }

        var firstRow = rows[0];
        sb.Append('|');
        foreach (var cell in firstRow.Elements<D.TableCell>())
        {
            var cellContent = new StringBuilder();
            if (cell.TextBody is not null)
            {
                foreach (var p in cell.TextBody.Elements<D.Paragraph>())
                {
                    var pContent = new StringBuilder();
                    foreach (var child in p.ChildElements)
                    {
                        if (child is D.Run run)
                        {
                            pContent.Append(run.Text?.Text);
                        }
                        else if (child.NamespaceUri == MathNamespace && child.LocalName == "oMath")
                        {
                            pContent.Append(MathParser.ParseOfficeMath(child));
                        }
                    }

                    if (cellContent.Length > 0 && pContent.Length > 0)
                    {
                        cellContent.Append("<br>");
                    }

                    cellContent.Append(pContent);
                }
            }

            var text = cellContent.ToString().Trim();
            sb.Append(' ').Append(text.Replace("|", "\\|", StringComparison.Ordinal)).Append(" |");
        }

        sb.AppendLine();

        sb.Append('|');

        foreach (var unused in firstRow.Elements<D.TableCell>())
        {
            sb.Append(" --- |");
        }

        sb.AppendLine();

        for (var i = 1; i < rows.Count; i++)
        {
            sb.Append('|');
            foreach (var cell in rows[i].Elements<D.TableCell>())
            {
                var cellContent = new StringBuilder();
                if (cell.TextBody is not null)
                {
                    foreach (var p in cell.TextBody.Elements<D.Paragraph>())
                    {
                        var pContent = new StringBuilder();
                        foreach (var child in p.ChildElements)
                        {
                            if (child is D.Run run)
                            {
                                pContent.Append(run.Text?.Text);
                            }
                            else if (child.NamespaceUri == MathNamespace && child.LocalName == "oMath")
                            {
                                pContent.Append(MathParser.ParseOfficeMath(child));
                            }
                        }

                        if (cellContent.Length > 0 && pContent.Length > 0)
                        {
                            cellContent.Append("<br>");
                        }

                        cellContent.Append(pContent);
                    }
                }

                var text = cellContent.ToString().Trim();
                sb.Append(' ').Append(text.Replace("|", "\\|", StringComparison.Ordinal)).Append(" |");
            }

            sb.AppendLine();
        }

        sb.AppendLine();
    }
}
