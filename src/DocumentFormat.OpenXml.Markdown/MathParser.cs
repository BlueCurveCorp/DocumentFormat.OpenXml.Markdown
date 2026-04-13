using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentFormat.OpenXml.Markdown;

/// <summary>
/// Internal parser to convert Office Math elements to LaTeX Markdown.
/// </summary>
internal static class MathParser
{
    private static readonly Dictionary<char, string> SymbolMap = new()
    {
        // Greek letters (Lowercase)
        { 'α', "\\alpha" }, { 'β', "\\beta" }, { 'γ', "\\gamma" }, { 'δ', "\\delta" },
        { 'ε', "\\epsilon" }, { 'ζ', "\\zeta" }, { 'η', "\\eta" }, { 'θ', "\\theta" },
        { 'ι', "\\iota" }, { 'κ', "\\kappa" }, { 'λ', "\\lambda" }, { 'μ', "\\mu" },
        { 'ν', "\\nu" }, { 'ξ', "\\xi" }, { 'ο', "o" }, { 'π', "\\pi" },
        { 'ρ', "\\rho" }, { 'σ', "\\sigma" }, { 'τ', "\\tau" }, { 'υ', "\\upsilon" },
        { 'φ', "\\phi" }, { 'χ', "\\chi" }, { 'ψ', "\\psi" }, { 'ω', "\\omega" },

        // Greek letters (Uppercase)
        { 'Α', "A" }, { 'Β', "B" }, { 'Γ', "\\Gamma" }, { 'Δ', "\\Delta" },
        { 'Ε', "E" }, { 'Ζ', "Z" }, { 'Η', "H" }, { 'Θ', "\\Theta" },
        { 'Ι', "I" }, { 'Κ', "K" }, { 'Λ', "\\Lambda" }, { 'Μ', "M" },
        { 'Ν', "N" }, { 'Ξ', "\\Xi" }, { 'Ο', "O" }, { 'Π', "\\Pi" },
        { 'Ρ', "P" }, { 'Σ', "\\Sigma" }, { 'Τ', "T" }, { 'Υ', "\\Upsilon" },
        { 'Φ', "\\Phi" }, { 'Χ', "X" }, { 'Ψ', "\\Psi" }, { 'Ω', "\\Omega" },

        // Operators and Symbols
        { '±', "\\pm" }, { '∓', "\\mp" }, { '×', "\\times" }, { '÷', "\\div" },
        { '⋅', "\\cdot" }, { '∞', "\\infty" }, { '∀', "\\forall" }, { '∃', "\\exists" },
        { '∈', "\\in" }, { '∉', "\\notin" }, { '∑', "\\sum" }, { '∏', "\\prod" },
        { '∫', "\\int" }, { '∬', "\\iint" }, { '∭', "\\iiint" }, { '∮', "\\oint" },
        { '√', "\\sqrt" }, { '∝', "\\propto" }, { '∠', "\\angle" }, { '∧', "\\wedge" },
        { '∨', "\\vee" }, { '∩', "\\cap" }, { '∪', "\\cup" }, { '≈', "\\approx" },
        { '≠', "\\neq" }, { '≤', "\\leq" }, { '≥', "\\geq" }, { '⊂', "\\subset" },
        { '⊃', "\\supset" }, { '⊆', "\\subseteq" }, { '⊇', "\\supseteq" },
        { '⇒', "\\Rightarrow" }, { '⇐', "\\Leftarrow" }, { '⇔', "\\Leftrightarrow" },
        { '→', "\\rightarrow" }, { '←', "\\leftarrow" }, { '↔', "\\leftrightarrow" },
        { '∂', "\\partial" }, { '∇', "\\nabla" }, { '∆', "\\Delta" }, { '…', "\\dots" },
        { '′', "'" }, { '″', "''" }, { '≅', "\\cong" }, { '≡', "\\equiv" },
        { '⊕', "\\oplus" }, { '⊗', "\\otimes" }, { '⊙', "\\odot" }, { '⊥', "\\perp" },
        { '⋄', "\\diamond" }, { '⋆', "\\star" },
    };

    private static readonly HashSet<string> CommonFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "sin", "cos", "tan", "csc", "sec", "cot",
        "arcsin", "arccos", "arctan",
        "sinh", "cosh", "tanh", "coth",
        "log", "ln", "lg", "lim", "max", "min", "sup", "inf", "det", "arg", "dim", "gcd", "hom", "ker"
    };

    public static string ParseOfficeMath(OpenXmlElement math)
    {
        if (math is null)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var child in math.ChildElements)
        {
            sb.Append(ToLatex(child));
        }

        var content = sb.ToString().Trim();
        return string.IsNullOrEmpty(content) ? string.Empty : "$" + content + "$";
    }

    public static string ParseOfficeMathPara(OpenXmlElement mathPara)
    {
        if (mathPara is null)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var math in mathPara.ChildElements.Where(x => x.LocalName == "oMath"))
        {
            var content = ParseOfficeMath(math);
            if (!string.IsNullOrEmpty(content))
            {
                sb.AppendLine(content);
                sb.AppendLine();
            }
        }

        return sb.ToString();
    }

    private static string ToLatex(OpenXmlElement element)
    {
        if (element is null)
        {
            return string.Empty;
        }

        if (element.LocalName == "t")
        {
            return MapText(element.InnerText);
        }

        return element.LocalName switch
        {
            "r" => GetContent(element),
            "f" => "\\frac{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "num")) + "}{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "den")) + "}",
            "rad" => string.IsNullOrEmpty(GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "deg")))
                               ? "\\sqrt{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")) + "}"
                               : "\\sqrt[" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "deg")) + "]{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")) + "}",
            "sSup" => "{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")) + "}^{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sup")) + "}",
            "sSub" => "{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")) + "}_{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sub")) + "}",
            "sSubSup" => "{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")) + "}_{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sub")) + "}^{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sup")) + "}",
            "sPre" => "_{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sub")) + "}^{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sup")) + "}{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")) + "}",
            "d" => "\\left( " + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")) + " \\right)",
            "nary" => "\\int_{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sub")) + "}^{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "sup")) + "} " + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e")),
            "m" => "\\begin{matrix} " + ParseMatrix(element) + " \\end{matrix}",
            "func" => ToFunctionName(element.ChildElements.FirstOrDefault(x => x.LocalName == "fName")) + " " + WrapContent(GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName == "e"))),
            "limLow" => ParseLimit(element, true),
            "limUpp" => ParseLimit(element, false),
            _ => GetContent(element),
        };
    }

    private static string GetContent(OpenXmlElement? element)
    {
        if (element is null)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var child in element.ChildElements)
        {
            sb.Append(ToLatex(child));
        }

        return sb.ToString();
    }

    private static string MapText(string text)
    {
        var sb = new StringBuilder();
        foreach (var c in text)
        {
            if (SymbolMap.TryGetValue(c, out var latex))
            {
                sb.Append(latex).Append(' ');
            }
            else
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }

    private static string ToFunctionName(OpenXmlElement? fName)
    {
        if (fName is null)
        {
            return string.Empty;
        }

        var content = GetContent(fName).Trim();
        if (string.IsNullOrEmpty(content))
        {
            return string.Empty;
        }

        if (content.StartsWith('\\'))
        {
            return content;
        }

        if (CommonFunctions.Contains(content))
        {
            return "\\" + content;
        }

        return "\\text{" + content + "}";
    }

    private static string WrapContent(string content)
    {
        if (string.IsNullOrEmpty(content))
        {
            return string.Empty;
        }

        if ((content.StartsWith('(') && content.EndsWith(')')) ||
            (content.StartsWith('{') && content.EndsWith('}')) ||
            (content.StartsWith('[') && content.EndsWith(']')))
        {
            return content;
        }

        return "{" + content + "}";
    }

    private static string ParseLimit(OpenXmlElement limitElement, bool isLower)
    {
        var e = limitElement.ChildElements.FirstOrDefault(x => x.LocalName == "e");
        var lim = limitElement.ChildElements.FirstOrDefault(x => x.LocalName == "lim");

        var eContent = GetContent(e).Trim();
        var limContent = GetContent(lim).Trim();

        if (eContent.Equals("lim", StringComparison.OrdinalIgnoreCase))
        {
            return (isLower ? "\\lim_{" : "\\overline{lim}_{") + limContent + "}";
        }

        return "{" + eContent + "}" + (isLower ? "_{" : "^{") + limContent + "}";
    }

    private static string ParseMatrix(OpenXmlElement matrix)
    {
        var sb = new StringBuilder();
        var rows = matrix.ChildElements.Where(x => x.LocalName == "mr").ToList();

        for (var i = 0; i < rows.Count; i++)
        {
            var args = rows[i].ChildElements.Where(x => x.LocalName == "arg").ToList();
            for (var j = 0; j < args.Count; j++)
            {
                sb.Append(GetContent(args[j]));
                if (j < args.Count - 1)
                {
                    sb.Append(" & ");
                }
            }

            if (i < rows.Count - 1)
            {
                sb.Append(" \\\\ ");
            }
        }

        return sb.ToString();
    }
}
