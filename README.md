# DocumentFormat.OpenXml.Markdown

A lightweight, powerful extension library for the official [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). It reliably converts Word (`.docx`), Excel (`.xlsx`), and PowerPoint (`.pptx`) documents into clean, GitHub-Flavored Markdown.

Perfect for migrating internal documents, formatting wiki content, or pre-processing Office files into LLM-friendly plain text!

## Features

- 📝 **Wordprocessing (`.docx`)**  
  Maps paragraph styles to standard Markdown headings. Parses inline text formatting like bold (`**`), italics (`*`), and strikethrough (`~~`). Supports **multi-level nested lists** and extracts hyperlinks and tables.
- 📊 **Spreadsheet (`.xlsx`)**  
  Resolves the SharedString table and extracts cached formula values. Renders worksheets as elegant GitHub-Flavored Markdown (GFM) tables.
- 📽️ **Presentation (`.pptx`)**  
  Iterates through slides sequentially. Re-maps slide titles to headings, and extracts nested text from Shapes, TextBoxes, and **Tables**.
- 🖼️ **Flexible Image Handling**  
  Easily control how embedded images are handled. You can safely ignore them, inline them directly via Base64 Data URIs, or export them to a local asset directory with custom URL prefixes.

## Getting Started

### Installation

*Note: Instructions will be finalized when the package is published to NuGet.*

```shell
dotnet add package DocumentFormat.OpenXml.Markdown
```

### Basic Usage

Usage is incredibly simple and stream-aware. You can pass raw OpenXML types or let the library infer the document naturally via a standard `Stream`.

#### Convert from a Stream

The converter will inspect the package's contents and automatically identify if it's Word, Excel, or PowerPoint.

```csharp
using DocumentFormat.OpenXml.Markdown;
using System.IO;

var fileStream = File.OpenRead("SampleDocument.docx");
string markdown = MarkdownConverter.ConvertToMarkdown(fileStream);

Console.WriteLine(markdown);
```

#### Convert from OpenXML objects

If you are already interacting with the Open XML SDK, you can convert specific documents you've already opened.

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Markdown;

using (var wordDoc = WordprocessingDocument.Open("SampleDocument.docx", false))
{
    string markdown = MarkdownConverter.ConvertToMarkdown(wordDoc);
}
```

### Advanced Usage: Image Handling Configurations

By default, any visual image found in the document stream is encoded internally as a Base64 string. However, you can change this behavior by instantiating `MarkdownConverterSettings`:

```csharp
var settings = new MarkdownConverterSettings 
{
    // Mode: Ignore         -> Ignores images entirely
    // Mode: Base64         -> Default behavior (![alt](data:image/png;base64,...))
    // Mode: ExportToFolder -> Writes images to disk and generates links
    ImageExportMode = ImageExportMode.ExportToFolder,

    // If using ExportToFolder, define where to persist the binaries:
    AssetExportDirectory = "./docs/assets/",

    // (Optional) Define a prefix for generated URLs (e.g. for CDN hosting):
    AssetLinkUrlPrefix = "https://cdn.example.com/assets/"
};

using var wordDoc = WordprocessingDocument.Open("Report.docx", false);
string markdown = MarkdownConverter.ConvertToMarkdown(wordDoc, settings);
```

### Complex Formatting Notes

- **Tables**: Small, simple tables work best. For Word and PowerPoint tables containing lists or multiple paragraphs in a single cell, the library uses `<br>` tags to preserve visual structure within GFM-compliant rows.
- **Lists**: Multi-level bullets are supported for both Word and PowerPoint, utilizing standard Markdown indentation logic.

## Office Math to LaTeX Supported Expressions

The `DocumentFormat.OpenXml.Markdown` parser comes with exhaustive support for Office Math (OMML) conversion to high-fidelity LaTeX within Markdown output.

It recursively searches through Word (`.docx`) nested structures (including tables) and PowerPoint (`.pptx`) DrawingML containers to identify and seamlessly extract mathematical structures.

## Supported OMML Mathematical Expressions

### 1. Basic Structures

- **Fractions (`f`)**: Automatically translated to `\frac{numerator}{denominator}` or `\binom{n}{k}` depending on the fractional bar presence.
- **Radicals (`rad`)**: N-th roots translated accurately (e.g. `\sqrt[n]{x}` vs square root `\sqrt{x}`).
- **Scripts & Indices**: Subscripts (`sSub`), Superscripts (`sSup`), Sub-Superscripts (`sSubSup`), and Left Sub-Superscripts/Pre-scripts (`sPre`) (e.g., `_{n}^{m}Y`).

### 2. Groupings & Delimiters
* **Delimiters (`d`)**: Automatically detects and translates bracket mappings (parentheses, braces `\{`, brackets `[`, absolute values `|`, norms `\|`, and floor/ceiling). Dynamically scales using `\left` and `\right`.
- **Delimiters (`d`)**: Automatically detects and translates bracket mappings (parentheses, braces `\{`, brackets `[`, absolute values `|`, norms `\|`, and floor/ceiling). Dynamically scales using `\left` and `\right`. 
- **Piecewise Functions (`cases`)**: Seamlessly identifies piecewise equations inside single-sided bracket structures `{` and maps them correctly to `\begin{cases} ... \end{cases}`.

### 3. Advanced Structures

- **Large Operators (`nary`)**: Supports Integrals (`\int`), Double/Triple Integrals, Summations (`\sum`), Products (`\prod`), and Contour Integrals, along with their respective lower bounds and upper limits.
- **Limiting Functions (`limLow`, `limUpp`)**: Detects layout instructions for generic functions (like `\lim`, `\max`, `\sup`) and properly applies subscript bounds via `\lim_{x \to \infty}` formatting.
- **Accents (`acc`)**: Replaces math accents correctly: `\hat`, `\vec`, `\tilde`, `\bar`, `\dot`, `\ddot`, `\breve`, etc.
- **Arrays & Matrices (`eqArr`, `m`)**: Maps equation alignments (`eqArr`) into `\begin{aligned}` systems, and native multi-dimensional cell arrays (`m`) into `\begin{matrix}` structures.
- **Group Characters (`groupChr`)**: Supports top-level and bottom-level layout formatting for grouping structures mapping to mathematical `\overset` or `\underset`.
- **Boxes and Underlines (`box`, `borderBox`, `bar`)**: Properly encapsulates equations inside `\boxed{}` and applies overlines/underlines (`\overline`, `\underline`) across structures.

### 4. Comprehensive Symbol Dictionary

Includes an extensive multi-character mapping dictionary designed to preserve visual parity during translation:

- All Uppercase and Lowercase Greek Symbols (`\alpha`, `\beta` ... `\Omega`)
- Set representations (`\in`, `\subset`, `\cup`, `\cap`)
- Propositional logic & arrows (`\forall`, `\exists`, `\rightarrow`, `\Leftrightarrow`)
- Math operations and relations (`\pm`, `\times`, `\leq`, `\neq`, `\approx`, `\equiv`)
- Function representations natively wrapped (e.g., trigonometric `\sin`, `\cos` and logarithmic `\ln`)

## Contributing

We welcome contributions! To collaborate on this repository:

1. Ensure you agree to the `.editorconfig` formatting rules.
2. Verify all modifications still allow for cross-compilation across .NET 10+ properly.
3. Write sufficient xUnit coverage for robust edge cases.

## License

This project is licensed under the same [MIT License](LICENSE) constraints as the official DocumentFormat.OpenXml SDK repository.
