using System;
using System.IO;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Markdown.Tests;

public class MarkdownConverterTests
{
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Test]
    public async Task WordDocument_SimpleParagraphs_ConvertedCorrectly()
    {
        // Arrange
        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());

            var p1 = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Hello World!")));
            var p2 = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId() { Val = "Heading1" }),
                new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Heading 1")));

            mainPart.Document.Body!.AppendChild(p1);
            mainPart.Document.Body!.AppendChild(p2);
        }

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("Hello World!");
        await Assert.That(markdown).Contains("# Heading 1");
    }

    [Test]
    public async Task WordDocument_Equation_ConvertedCorrectly()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var officeMath = new DocumentFormat.OpenXml.Math.OfficeMath();
            officeMath.AppendChild(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("x")));
            officeMath.AppendChild(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("^")));
            officeMath.AppendChild(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("2")));

            var mathPara = new DocumentFormat.OpenXml.Math.Paragraph(officeMath);
            mainPart.Document.Body!.Append(mathPara);
        }

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        await Assert.That(markdown).Contains("$x^2$");
    }

    [Test]
    public async Task WordDocument_Equation_Strict()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\equation.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains(@"$\sin {\alpha }\pm \sin {\beta }=2\sin {\frac{1}{2}\left( \alpha \pm \beta  \right)}\cos {\frac{1}{2}\left( \alpha \mp \beta  \right)}$");
        await Assert.That(markdown).Contains(@"$\lim_{n\rightarrow \infty} {\left( 1+\frac{1}{n} \right)}^{n}$");
        await Assert.That(markdown).Contains(@"${e}^{x}=1+\frac{x}{1!}+\frac{{x}^{2}}{2!}+\frac{{x}^{3}}{3!}+\dots ,  -\infty <x<\infty$");
    }


    [Test]
    public async Task WordDocument_Equation_Transitional()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\equation.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains(@"$\sin {\alpha }\pm \sin {\beta }=2\sin {\frac{1}{2}\left( \alpha \pm \beta  \right)}\cos {\frac{1}{2}\left( \alpha \mp \beta  \right)}$");
        await Assert.That(markdown).Contains(@"$\lim_{n\rightarrow \infty} {\left( 1+\frac{1}{n} \right)}^{n}$");
        await Assert.That(markdown).Contains(@"${e}^{x}=1+\frac{x}{1!}+\frac{{x}^{2}}{2!}+\frac{{x}^{3}}{3!}+\dots ,  -\infty <x<\infty$");
    }

    [Test]
    public async Task WordDocument_Strict_Workflow()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\WORKFLOW_ENGINE_SPEC.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("Interopérabilité");
        await Assert.That(markdown).Contains("PhotoTrailer");
    }

    [Test]
    public async Task WordDocument_Transitional_Workflow()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\WORKFLOW_ENGINE_SPEC.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("Interopérabilité");
        await Assert.That(markdown).Contains("PhotoTrailer");
    }

    [Test]
    public async Task WordDocument_Transitional_Large()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\15-MB-docx-file-download.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("Suspendisse molestie nibh magna, eu maximus ante tincidunt tincidunt. Ut sit amet bibendum turpis, vitae iaculis nibh");
    }

    [Test]
    public async Task WordDocument_Strict_Large()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\15-MB-docx-file-download.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("Suspendisse molestie nibh magna, eu maximus ante tincidunt tincidunt. Ut sit amet bibendum turpis, vitae iaculis nibh");
    }

    [Test]
    public async Task WordDocument_OpenXml_Strict()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\Open XML SDK.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains(".0.0 introduced breaking changes and you should be able");
        await Assert.That(markdown).Contains("Packages");
        await Assert.That(markdown).Contains("DocumentFormat.OpenXml.Framework");
        await Assert.That(markdown).Contains("Prerelease");
        await Assert.That(markdown).Contains("Spreadsheet Samples");
        await Assert.That(markdown).Contains("**Open XML SDK 2.5 Productivity Tool**");
    }

    [Test]
    public async Task WordDocument_OpenXml_Transitional()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\Open XML SDK.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains(".0.0 introduced breaking changes and you should be able");
        await Assert.That(markdown).Contains("Packages");
        await Assert.That(markdown).Contains("DocumentFormat.OpenXml.Framework");
        await Assert.That(markdown).Contains("Prerelease");
        await Assert.That(markdown).Contains("Spreadsheet Samples");
        await Assert.That(markdown).Contains("**Open XML SDK 2.5 Productivity Tool**");
    }

    [Test]
    public async Task WordDocument_Cv_Demo_Transitional()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\cv_demo.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("Assistant juridique The Phone Company  Résumez vos responsabilités et réalisations clés");
        await Assert.That(markdown).Contains("Indiquez vos objectifs professionnels et montrez leur alignement avec la description de la tâche que vous ciblez");
    }

    [Test]
    public async Task WordDocument_Cv_Demo_Strict()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\cv_demo.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("Assistant juridique The Phone Company  Résumez vos responsabilités et réalisations clés");
        await Assert.That(markdown).Contains("Indiquez vos objectifs professionnels et montrez leur alignement avec la description de la tâche que vous ciblez");
    }


    [Test]
    public async Task WordDocument_Tuto_Strict()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\tuto.docx");

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("- Dans l’onglet Références, cliquez sur Table des matières, puis dans la partie inférieure, cliquez sur Table des matières personnalisée.");
        await Assert.That(markdown).Contains("Créer, mettre à jour et personnaliser une table des matières");
        await Assert.That(markdown).Contains("Accédez à Rechercher des outils adaptés dans la partie supérieure de la fenêtre, puis entrez ce que vous voulez faire.");
        await Assert.That(markdown).Contains("| Table H1 | Table H2 | Table H3 | Table H4 |");
        await Assert.That(markdown).Contains("| --- | --- | --- | --- |");
        await Assert.That(markdown).Contains("| A | AA | AZ | AE |");
        await Assert.That(markdown).Contains("| B | BB | BZ | BE |");
        await Assert.That(markdown).Contains("| C | CC | CZ | CE |");
        await Assert.That(markdown).Contains("| D | - Table bullet 1 - Table bullet 2 |  |  |");
        await Assert.That(markdown).Contains("- Bullet 1");
        await Assert.That(markdown).Contains("- Bullet 2");
        await Assert.That(markdown).Contains("- Bullet 3");
        await Assert.That(markdown).Contains("- Number 1");
        await Assert.That(markdown).Contains("- Number 2");
        await Assert.That(markdown).Contains("- Number 3");
        await Assert.That(markdown).Contains("- Hierachy 1");
        await Assert.That(markdown).Contains("- Hierarchy 2");
        await Assert.That(markdown).Contains("- Hiearachy 3");
        await Assert.That(markdown).Contains("- Hierarchy 2");
        await Assert.That(markdown).Contains("- Apple");
        await Assert.That(markdown).Contains("- Pear");
        await Assert.That(markdown).Contains("- Fruit");
    }


    [Test]
    public async Task Excel_SimpleTable_ConvertedCorrectly()
    {
        // Arrange
        using var stream = new MemoryStream();
        using (var spreadsheetDoc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = spreadsheetDoc.AddWorkbookPart();
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

            var sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
            var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "TestSheet" };
            sheets.Append(sheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!;
            var row1 = new DocumentFormat.OpenXml.Spreadsheet.Row();
            row1.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Header 1"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String, CellReference = "A1" });
            row1.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Header 2"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String, CellReference = "B1" });
            sheetData.Append(row1);

            var row2 = new DocumentFormat.OpenXml.Spreadsheet.Row();
            row2.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Value 1"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String, CellReference = "A2" });
            row2.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell { CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Value 2"), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String, CellReference = "B2" });
            sheetData.Append(row2);
        }


        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## TestSheet");
        await Assert.That(markdown).Contains("| Header 1 | Header 2 |");
        await Assert.That(markdown).Contains("| Value 1 | Value 2 |");
        await Assert.That(markdown).Contains("| --- | --- |");
    }


    [Test]
    public async Task Excel_XLSX_5000_Strict()
    {
        // Arrange
        // Arrange
        using var stream = File.OpenRead(@"data\strict\file_example_XLSX_5000.xlsx");

        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## Sheet1");
        await Assert.That(markdown).Contains("|  | First Name | Last Name | Gender | Country | Age | Date | Id |");
        await Assert.That(markdown).Contains("| 47 | Felisa | Cail | Female | United States | 28 | 16/08/2016 | 6525 |");
        await Assert.That(markdown).Contains("| 3418 | Lauralee | Perrine | Female | Great Britain | 28 | 16/08/2016 | 6597 |");
    }

    [Test]
    public async Task Excel_XLSX_5000_Transitional()
    {
        // Arrange
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\file_example_XLSX_5000.xlsx");
        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## Sheet1");
        await Assert.That(markdown).Contains("|  | First Name | Last Name | Gender | Country | Age | Date | Id |");
        await Assert.That(markdown).Contains("| 47 | Felisa | Cail | Female | United States | 28 | 16/08/2016 | 6525 |");
        await Assert.That(markdown).Contains("| 3418 | Lauralee | Perrine | Female | Great Britain | 28 | 16/08/2016 | 6597 |");
    }



    [Test]
    public async Task Excel_Combining_Lists_Strict()
    {
        // Arrange
        // Arrange
        using var stream = File.OpenRead(@"data\strict\Combining-Lists.xlsx");
        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## Combine Lists");
        await Assert.That(markdown).Contains("| Combining Lists |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |");
        await Assert.That(markdown).Contains("| Sheffield United |  | Nott'm Forest |  | Northampton |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  |  |  |  |  |  |");
    }

    [Test]
    public async Task Excel_Combining_Lists_Transitional()
    {
        // Arrange
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\Combining-Lists.xlsx");
        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## Combine Lists");
        await Assert.That(markdown).Contains("| Combining Lists |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |");
        await Assert.That(markdown).Contains("| Sheffield United |  | Nott'm Forest |  | Northampton |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  | Sheffield United |  |  |  |  |  |  |");
    }



    [Test]
    public async Task Excel_League_Table_Transitional()
    {
        // Arrange
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\League-Table-Examples.xlsx");
        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## Part B");
        await Assert.That(markdown).Contains("## Part A");
        await Assert.That(markdown).Contains("## Part C");
        await Assert.That(markdown).Contains("## Data");
        await Assert.That(markdown).Contains("| Part B |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | 3 | Bournemouth | 38 | 9 | 7 | 22 | 40 | 65 | -25 | 34 |  |  |  | 3 | Manchester United | 38 | 18 | 12 | 8 | 66 | 36 | 30 | 66 |  |  |  | 3 | Manchester United | 38 | 18 | 12 | 8 | 66 | 36 | 30 | 66 |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("| Part A |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | 14 | 7 | Crystal Palace | 38 | 11 | 10 | 17 | 31 | 50 | -19 | 42.999813093229996 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59.000115091100007 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59.000115091100007 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59.000115091100007 |  |");
        await Assert.That(markdown).Contains("| team_home | team_away | home_goal | away_goal | played | season | date_time | Result |");
        await Assert.That(markdown).Contains("| --- | --- | --- | --- | --- | --- | --- | --- |");
        await Assert.That(markdown).Contains("| Liverpool | Norwich City | 4 | 1 | 1 | 2019/20 | 43686.833333333336 | H |");
        await Assert.That(markdown).Contains("| West Ham United | Manchester City | 0 | 5 | 1 | 2019/20 | 43687.520833333336 | A |");
        await Assert.That(markdown).Contains("| Part C |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  | Table C1 (Ordered) |  |  |  |  |  |  |  |  |  |  |  |  | Table C2 (Ordered) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | POS | TEAM | P | W | D | L | F | A | GD | PTS |  |  |  | POS | TEAM | P | W | D | L | F | A | GD | PTS |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59 |  |  |  |  |  |  |  |  |  |  |  |");
    }

    [Test]
    public async Task Excel_League_Table_Strict()
    {
        // Arrange
        // Arrange
        using var stream = File.OpenRead(@"data\strict\League-Table-Examples.xlsx");
        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## Part B");
        await Assert.That(markdown).Contains("## Part A");
        await Assert.That(markdown).Contains("## Part C");
        await Assert.That(markdown).Contains("## Data");
        await Assert.That(markdown).Contains("| Part B |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | 3 | Bournemouth | 38 | 9 | 7 | 22 | 40 | 65 | -25 | 34 |  |  |  | 3 | Manchester United | 38 | 18 | 12 | 8 | 66 | 36 | 30 | 66 |  |  |  | 3 | Manchester United | 38 | 18 | 12 | 8 | 66 | 36 | 30 | 66 |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("| Part A |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | 14 | 7 | Crystal Palace | 38 | 11 | 10 | 17 | 31 | 50 | -19 | 42.999813093229996 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59.000115091100007 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59.000115091100007 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59.000115091100007 |  |");
        await Assert.That(markdown).Contains("| team_home | team_away | home_goal | away_goal | played | season | date_time | Result |");
        await Assert.That(markdown).Contains("| --- | --- | --- | --- | --- | --- | --- | --- |");
        await Assert.That(markdown).Contains("| Liverpool | Norwich City | 4 | 1 | 1 | 2019/20 | 2019-08-09T20:00:00.00000020954757600 | H |");
        await Assert.That(markdown).Contains("| West Ham United | Manchester City | 0 | 5 | 1 | 2019/20 | 2019-08-10T12:30:00.00000020954757600 | A |");
        await Assert.That(markdown).Contains("| Part C |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  | Table C1 (Ordered) |  |  |  |  |  |  |  |  |  |  |  |  | Table C2 (Ordered) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | POS | TEAM | P | W | D | L | F | A | GD | PTS |  |  |  | POS | TEAM | P | W | D | L | F | A | GD | PTS |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59 |  |  |  | 7 | Wolverhampton Wanderers | 38 | 15 | 14 | 9 | 51 | 40 | 11 | 59 |  |  |  |  |  |  |  |  |  |  |  |");
    }

    [Test]
    public async Task Excel_Xlsm()
    {
        // Arrange
        // Arrange
        using var stream = File.OpenRead(@"data\strict\How-Excel-Stores-and-Displays-Data.xlsm");
        stream.Position = 0;
        using var readDoc = SpreadsheetDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("|  |  |  |  |  |  |  | Acknowledgment | 6 | ACK | 65 | 6 | _x0006_ |  | 7 | _x0007_ |  | 7 | 6 | _x0006_ |  | 1 > A | 119 > 131 | 49 > 65 | 0 |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("## Part 2 Extras");
        await Assert.That(markdown).Contains("|  |  |  |  |  |  |  | Table P2E.1 |  |  |  |  |  |  | Table P2E.2");
        await Assert.That(markdown).Contains("|  |  |  |  |  |  |  | Uppercase V | 86 | V | 86 | 86 | V |  | 87 | W |  | 87 | 128 | € |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("|  |  |  |  |  |  |  | Custom Formatted (;;;) | 0 | 0 |  |  |  |  |  |  |  |  |  |");
        await Assert.That(markdown).Contains("## Part 1");
    }


    [Test]
    public async Task PowerPoint_SimpleSlide_ConvertedCorrectly()
    {
        // Arrange
        using var stream = new MemoryStream();
        using (var presentationDoc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();
            var slideMasterIdList = new DocumentFormat.OpenXml.Presentation.SlideMasterIdList(new DocumentFormat.OpenXml.Presentation.SlideMasterId() { Id = 2147483648U, RelationshipId = "rId1" });
            var slideIdList = new DocumentFormat.OpenXml.Presentation.SlideIdList(new DocumentFormat.OpenXml.Presentation.SlideId() { Id = 256U, RelationshipId = "rId2" });
            presentationPart.Presentation.Append(slideMasterIdList, slideIdList);

            var slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart.Slide = new DocumentFormat.OpenXml.Presentation.Slide(new DocumentFormat.OpenXml.Presentation.CommonSlideData(new DocumentFormat.OpenXml.Presentation.ShapeTree()));

            var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

            // Add a title shape
            var titleShape = new DocumentFormat.OpenXml.Presentation.Shape();
            titleShape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 1U, Name = "Title" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties(new DocumentFormat.OpenXml.Presentation.PlaceholderShape { Type = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title }));

            titleShape.TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                new DocumentFormat.OpenXml.Drawing.ListStyle(),
                new DocumentFormat.OpenXml.Drawing.Paragraph(new DocumentFormat.OpenXml.Drawing.Run(new DocumentFormat.OpenXml.Drawing.Text("Slide Title"))));
            shapeTree.Append(titleShape);

            // Add a content shape
            var contentShape = new DocumentFormat.OpenXml.Presentation.Shape();
            contentShape.NonVisualShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 2U, Name = "Content" },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

            contentShape.TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                new DocumentFormat.OpenXml.Drawing.ListStyle(),
                new DocumentFormat.OpenXml.Drawing.Paragraph(new DocumentFormat.OpenXml.Drawing.Run(new DocumentFormat.OpenXml.Drawing.Text("Hello from PPT!"))));
            shapeTree.Append(contentShape);
        }


        stream.Position = 0;
        using var readDoc = PresentationDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## Slide Title");
        await Assert.That(markdown).Contains("Hello from PPT!");
    }

    [Test]
    public async Task PowerPoint_TechTrend_Transitional()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\transitional\Tech Trends 2026 by Slidesgo.pptx");
        stream.Position = 0;
        using var readDoc = PresentationDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## TECH TRENDS ");
        await Assert.That(markdown).Contains("To view this template correctly in PowerPoint, download and install the fonts we used");
        await Assert.That(markdown).Contains("## CONTENTS OF THIS TEMPLATE");
        await Assert.That(markdown).Contains("- Do you know what helps you make your point clear?Lists like this one:");
        await Assert.That(markdown).Contains("- They’re simple ");
        await Assert.That(markdown).Contains("- You’ll never forget to buy milk!");
        await Assert.That(markdown).Contains("This a text zone content");
        await Assert.That(markdown).Contains("- Bullet 1");
        await Assert.That(markdown).Contains("Simple word ART");
        await Assert.That(markdown).Contains("| Table H1 | Table H2 | Table H3 | Table H4 |");
        await Assert.That(markdown).Contains("| --- | --- | --- | --- |");
        await Assert.That(markdown).Contains("| A | AA | AZ | AE |");
        await Assert.That(markdown).Contains("| B | BB | BZ | BE |");
        await Assert.That(markdown).Contains("| C | CC | CZ | CE |");
    }

    [Test]
    public async Task PowerPoint_TechTrend_Strict()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\Tech Trends 2026 by Slidesgo.pptx");
        stream.Position = 0;
        using var readDoc = PresentationDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Assert
        await Assert.That(markdown).Contains("## TECH TRENDS ");
        await Assert.That(markdown).Contains("To view this template correctly in PowerPoint, download and install the fonts we used");
        await Assert.That(markdown).Contains("## CONTENTS OF THIS TEMPLATE");
        await Assert.That(markdown).Contains("- Do you know what helps you make your point clear?Lists like this one:");
        await Assert.That(markdown).Contains("- They’re simple ");
        await Assert.That(markdown).Contains("- You’ll never forget to buy milk!");
        await Assert.That(markdown).Contains("This a text zone content");
        await Assert.That(markdown).Contains("- Bullet 1");
        await Assert.That(markdown).Contains("Simple word ART");
        await Assert.That(markdown).Contains("| Table H1 | Table H2 | Table H3 | Table H4 |");
        await Assert.That(markdown).Contains("| --- | --- | --- | --- |");
        await Assert.That(markdown).Contains("| A | AA | AZ | AE |");
        await Assert.That(markdown).Contains("| B | BB | BZ | BE |");
        await Assert.That(markdown).Contains("| C | CC | CZ | CE |");
    }

    [Test]
    public void MarkdownConverterSettings_ExportToFolderWithoutDirectory_Throws()
    {
        // Arrange
        var settings = new MarkdownConverterSettings
        {
            ImageExportMode = ImageExportMode.ExportToFolder,
            AssetExportDirectory = " ", // Invalid
        };

        // Act & Assert
        Assert.Throws<InvalidOperationException>(settings.Validate);
    }

    [Test]
    public async Task WordDocument_ImageExportToFolder_ConvertedCorrectly()
    {
        // Arrange
        var tempPath = Path.Combine(Path.GetTempPath(), "OpenXmlMarkdownTests_" + Guid.NewGuid().ToString("n"));

        _ = Directory.CreateDirectory(tempPath);
        try
        {
            using var stream = new MemoryStream();
            using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());

                // Add an image part
                var imagePart = mainPart.AddImagePart(ImagePartType.Png, "rIdImage1");
                var imageBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }; // PNG header

                using (var imageStream = imagePart.GetStream())
                {
                    await imageStream.WriteAsync(imageBytes);
                }

                // Add a drawing element with the blip
                var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 990000L, Cy = 792000L },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = 1U, Name = "Picture 1" },
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() { Id = 0U, Name = "New Bitmap Image.png" },
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new DocumentFormat.OpenXml.Drawing.Blip() { Embed = "rIdImage1" },
                                        new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties()
                                )
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                );

                mainPart.Document.Body!.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(drawing)));
            }

            stream.Position = 0;
            using var readDoc = WordprocessingDocument.Open(stream, false);

            var settings = new MarkdownConverterSettings
            {
                ImageExportMode = ImageExportMode.ExportToFolder,
                AssetExportDirectory = tempPath,
                AssetLinkUrlPrefix = "https://cdn.example.com/assets/",
            };

            // Act
            var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc, settings);

            var files = Directory.GetFiles(tempPath);

            // Assert
            foreach (var file in files)
            {
                await Assert.That(markdown).Contains($"![Image](https://cdn.example.com/assets/{Path.GetFileName(file)})");
            }
            await Assert.That(files).HasAtMost(1);
            await Assert.That(files[0]).EndsWith(".png");
        }
        finally
        {
            if (Directory.Exists(tempPath))
            {
                Directory.Delete(tempPath, true);
            }
        }
    }

    [Test]
    public async Task PowerPoint_ImageExportToFolder_ConvertedCorrectly()
    {
        // Arrange
        var tempPath = Path.Combine(Path.GetTempPath(), "OpenXmlMarkdownTests_PPT_" + Guid.NewGuid().ToString("n"));

        Directory.CreateDirectory(tempPath);

        try
        {
            using var stream = new MemoryStream();
            using (var presentationDoc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
            {
                var presentationPart = presentationDoc.AddPresentationPart();
                presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();
                var slideIdList = new DocumentFormat.OpenXml.Presentation.SlideIdList(new DocumentFormat.OpenXml.Presentation.SlideId() { Id = 256U, RelationshipId = "rId1" });
                presentationPart.Presentation.Append(slideIdList);

                var slidePart = presentationPart.AddNewPart<SlidePart>("rId1");
                slidePart.Slide = new DocumentFormat.OpenXml.Presentation.Slide(new DocumentFormat.OpenXml.Presentation.CommonSlideData(new DocumentFormat.OpenXml.Presentation.ShapeTree()));

                // Add an image part
                var imagePart = slidePart.AddImagePart(ImagePartType.Png, "rIdImage1");
                var imageBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }; // PNG header

                using (var imageStream = imagePart.GetStream())
                {
                    await imageStream.WriteAsync(imageBytes);
                }

                // Add a picture to the slide
                var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;

                var pic = new DocumentFormat.OpenXml.Presentation.Picture(
                    new DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 4U, Name = "Picture 1" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                    new DocumentFormat.OpenXml.Presentation.BlipFill(
                        new DocumentFormat.OpenXml.Drawing.Blip() { Embed = "rIdImage1" },
                        new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                    new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                        new DocumentFormat.OpenXml.Drawing.Transform2D(
                            new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                            new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 990000L, Cy = 792000L }),
                        new DocumentFormat.OpenXml.Drawing.PresetGeometry(new DocumentFormat.OpenXml.Drawing.AdjustValueList()) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }));

                shapeTree.Append(pic);
            }

            stream.Position = 0;
            using var readDoc = PresentationDocument.Open(stream, false);

            var settings = new MarkdownConverterSettings
            {
                ImageExportMode = ImageExportMode.ExportToFolder,
                AssetExportDirectory = tempPath,
                AssetLinkUrlPrefix = "./",
            };

            // Act
            var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc, settings);
            var files = Directory.GetFiles(tempPath);
            // Assert

            foreach (var file in files)
            {
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
                var name = fileNameWithoutExtension[..fileNameWithoutExtension.LastIndexOf('_')];

                await Assert.That(markdown).Contains($"![{name}](./{Path.GetFileName(file)})");
            }

            await Assert.That(files).HasAtMost(1);
            await Assert.That(files[0]).EndsWith(".png");
        }
        finally
        {
            if (Directory.Exists(tempPath))
            {
                Directory.Delete(tempPath, true);
            }
        }
    }

    [Test]
    public async Task PowerPoint_Equation_ConvertedCorrectly()
    {
        using var stream = new MemoryStream();
        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation(new DocumentFormat.OpenXml.Presentation.SlideIdList(new DocumentFormat.OpenXml.Presentation.SlideId() { Id = 256, RelationshipId = "rId1" }));

            var slidePart = presentationPart.AddNewPart<SlidePart>("rId1");

            var math = new DocumentFormat.OpenXml.Math.OfficeMath();
            math.AppendChild(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("y")));
            math.AppendChild(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("=")));
            math.AppendChild(new DocumentFormat.OpenXml.Math.Run(new DocumentFormat.OpenXml.Math.Text("x")));

            slidePart.Slide = new DocumentFormat.OpenXml.Presentation.Slide(
                new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                    new DocumentFormat.OpenXml.Presentation.ShapeTree(
                        new DocumentFormat.OpenXml.Presentation.Shape(
                            new DocumentFormat.OpenXml.Presentation.TextBody(
                                new DocumentFormat.OpenXml.Drawing.Paragraph(math))))));
        }

        stream.Position = 0;
        using var readDoc = PresentationDocument.Open(stream, false);

        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        await Assert.That(markdown).Contains("$y=x$");
    }


    [Test]
    public async Task PowerPoint_Equation_Strict()
    {
        // Arrange
        using var stream = File.OpenRead(@"data\strict\equation.pptx");
        stream.Position = 0;
        using var readDoc = PresentationDocument.Open(stream, false);

        // Act
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);
 
        await Assert.That(markdown).Contains("$y=x$");
    }

    [Test]
    public async Task WordDocument_ExponentialSeries_ConvertedCorrectly()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Create: e^x = 1 + x/1! + ... , -inf < x < inf
            var officeMath = CreateMathElement("oMath");

            // e^x
            var sSup = CreateMathElement("sSup");
            var eBase = CreateMathElement("e");
            eBase.AppendChild(CreateMathRun("e"));
            var supVal = CreateMathElement("sup");
            supVal.AppendChild(CreateMathRun("x"));
            sSup.AppendChild(eBase);
            sSup.AppendChild(supVal);
            officeMath.AppendChild(sSup);

            officeMath.AppendChild(CreateMathRun("="));
            officeMath.AppendChild(CreateMathRun("1"));
            officeMath.AppendChild(CreateMathRun("+"));

            // x/1!
            var f = CreateMathElement("f");
            var num = CreateMathElement("num");
            num.AppendChild(CreateMathRun("x"));
            var den = CreateMathElement("den");
            den.AppendChild(CreateMathRun("1!"));
            f.AppendChild(num);
            f.AppendChild(den);
            officeMath.AppendChild(f);

            officeMath.AppendChild(CreateMathRun("+"));
            officeMath.AppendChild(CreateMathRun("\u2026")); // dots
            officeMath.AppendChild(CreateMathRun(","));
            officeMath.AppendChild(CreateMathRun("-"));
            officeMath.AppendChild(CreateMathRun("\u221E")); // inf
            officeMath.AppendChild(CreateMathRun("<"));
            officeMath.AppendChild(CreateMathRun("x"));
            officeMath.AppendChild(CreateMathRun("<"));
            officeMath.AppendChild(CreateMathRun("\u221E")); // inf

            var mathPara = CreateMathElement("oMathPara");
            mathPara.AppendChild(officeMath);
            mainPart.Document.Body!.Append(mathPara);
        }

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        await Assert.That(markdown).Contains(@"e^{x}");
        await Assert.That(markdown).Contains(@"\frac{x}{1!}");
        await Assert.That(markdown).Contains(@"\dots");
        await Assert.That(markdown).Contains(@"\infty");
    }


    [Test]
    public async Task WordDocument_LimitEquation_ConvertedCorrectly()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Create: lim_{n->inf} (1 + 1/n)^n
            var officeMath = new DocumentFormat.OpenXml.Math.OfficeMath();

            // limLow (e="lim", lim="n->inf")
            var limLow = CreateMathElement("limLow");
            var e = CreateMathElement("e");
            e.AppendChild(CreateMathRun("lim"));
            var lim = CreateMathElement("lim");
            lim.AppendChild(CreateMathRun("n\u2192\u221E")); // n->inf

            limLow.AppendChild(e);
            limLow.AppendChild(lim);

            // Sup (e="(1 + 1/n)", sup="n")
            var sup = CreateMathElement("sSup");
            var e2 = CreateMathElement("e");
            e2.AppendChild(CreateMathRun("("));
            e2.AppendChild(CreateMathRun("1"));
            e2.AppendChild(CreateMathRun("+"));
            var f = CreateMathElement("f");
            var num = CreateMathElement("num");
            num.AppendChild(CreateMathRun("1"));
            var den = CreateMathElement("den");
            den.AppendChild(CreateMathRun("n"));
            f.AppendChild(num);
            f.AppendChild(den);
            e2.AppendChild(f);
            e2.AppendChild(CreateMathRun(")"));

            var super = CreateMathElement("sup");
            super.AppendChild(CreateMathRun("n"));

            sup.AppendChild(e2);
            sup.AppendChild(super);

            officeMath.AppendChild(limLow);
            officeMath.AppendChild(sup);

            var mathPara = new DocumentFormat.OpenXml.Math.Paragraph(officeMath);
            mainPart.Document.Body!.Append(mathPara);
        }

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);
        var markdown = await MarkdownConverter.ConvertToMarkdownAsync(readDoc);

        // Actual output was: $\lim_{n\rightarrow \infty}{(1+\frac{1}{n})}^{n}$
        await Assert.That(markdown).Contains(@"\lim_{n\rightarrow \infty}");
        await Assert.That(markdown).Contains(@"\frac{1}{n}");
        await Assert.That(markdown).Contains(@"^{n}");

    }

    private static OpenXmlElement CreateMathElement(string tagName)
    {
        return new OpenXmlUnknownElement("m", tagName, MathNamespace);
    }

    private static OpenXmlElement CreateMathElementWithVal(string tagName, string val)
    {
        var element = CreateMathElement(tagName);
        element.SetAttribute(new OpenXmlAttribute(string.Empty, "val", string.Empty, val));
        return element;
    }

    private static OpenXmlElement CreateMathRun(string text)
    {
        var r = CreateMathElement("r");
        var t = CreateMathElement("t");
        t.AppendChild(new DocumentFormat.OpenXml.Math.Text(text));
        r.AppendChild(t);
        return r;
    }
}
