// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Internal parser to convert SpreadsheetDocument to Markdown.
/// </summary>
internal static class SpreadsheetParser
{
    [SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Respect interface")]
    public static string Parse(SpreadsheetDocument document, MarkdownConverterSettings settings)
    {
        var sb = new StringBuilder();
        var workbookPart = document.WorkbookPart;

        if (workbookPart is null)
        {
            return string.Empty;
        }

        var stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var stringTable = stringTablePart?.SharedStringTable;

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var sheetName = "Sheet";
            var sheetId = workbookPart.GetIdOfPart(worksheetPart);
            var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault(s => s.Id?.Value == sheetId);

            if (sheet?.Name?.Value is not null)
            {
                sheetName = sheet.Name.Value;
            }

            sb.AppendFormat(CultureInfo.InvariantCulture, "## {0}", sheetName).AppendLine();
            sb.AppendLine();

            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet?.Elements<SheetData>().FirstOrDefault();

            if (sheetData is not null)
            {
                var rows = sheetData.Elements<Row>().ToList();
                if (rows.Count > 0)
                {
                    var maxCol = 0;

                    foreach (var row in rows)
                    {
                        foreach (var cell in row.Elements<Cell>())
                        {
                            var colIndex = GetColumnIndex(cell.CellReference?.Value);

                            if (colIndex > maxCol)
                            {
                                maxCol = colIndex;
                            }
                        }
                    }

                    var isFirstRow = true;

                    foreach (var row in rows)
                    {
                        var cells = row.Elements<Cell>().ToDictionary(c => GetColumnIndex(c.CellReference?.Value));

                        sb.Append('|');
                        for (var i = 1; i <= maxCol; i++)
                        {
                            var cellValue = cells.TryGetValue(i, out var cell) ? GetCellValue(cell, stringTable) : string.Empty;

                            cellValue = cellValue.Replace("|", "\\|", StringComparison.Ordinal);

                            sb.AppendFormat(CultureInfo.InvariantCulture, " {0} |", cellValue);
                        }
                        sb.AppendLine();

                        if (isFirstRow)
                        {
                            sb.Append('|');
                            for (var i = 1; i <= maxCol; i++)
                            {
                                sb.Append(" --- |");
                            }
                            sb.AppendLine();
                            isFirstRow = false;
                        }
                    }
                }
            }

            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return 0;
        }

        var columnReference = new string([.. cellReference.TakeWhile(char.IsLetter)]);
        var columnIndex = 0;
        var factor = 1;

        for (var i = columnReference.Length - 1; i >= 0; i--)
        {
            columnIndex += (columnReference[i] - 'A' + 1) * factor;
            factor *= 26;
        }
        return columnIndex;
    }

    private static string GetCellValue(Cell cell, SharedStringTable? stringTable)
    {
        if (cell.CellValue is null)
        {
            return string.Empty;
        }

        var value = cell.CellValue.Text;

        if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
        {
            if (int.TryParse(value, out var ssid) && stringTable is not null)
            {
                var item = stringTable.ElementAtOrDefault(ssid);
                if (item is not null)
                {
                    return item.InnerText;
                }
            }
        }
        else if (cell.CellFormula is not null)
        {
            // If the cell contains a formula, the CellValue typically contains the cached computed value
            // We return that directly.
            return value;
        }

        return value;
    }
}
