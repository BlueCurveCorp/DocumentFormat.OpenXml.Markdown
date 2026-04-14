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
        { 'ϵ', "\\varepsilon" }, { 'ϑ', "\\vartheta" }, { 'ϖ', "\\varpi" }, { 'ϱ', "\\varrho" },
        { 'ς', "\\varsigma" }, { 'ϕ', "\\varphi" },

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
        { '⋄', "\\diamond" }, { '⋆', "\\star" }, { '∗', "\\ast" },
        { '∼', "\\sim" }, { '≃', "\\simeq" }, { '≍', "\\asymp" },
        { '≐', "\\doteq" }, { '≪', "\\ll" }, { '≫', "\\gg" }, { '≺', "\\prec" }, { '≻', "\\succ" },
        { '≼', "\\preccurlyeq" }, { '≽', "\\succcurlyeq" }, { '⊏', "\\sqsubset" }, { '⊐', "\\sqsupset" },
        { '⊑', "\\sqsubseteq" }, { '⊒', "\\sqsupseteq" }, { '⊢', "\\vdash" }, { '⊣', "\\dashv" },
        { '∐', "\\coprod" }, { '⋀', "\\bigwedge" }, { '⋁', "\\bigvee" }, { '⋂', "\\bigcap" },
        { '⋃', "\\bigcup" }, { '⊎', "\\uplus" }, { '∔', "\\dotplus" }, { '∖', "\\setminus" },
        { '∘', "\\circ" }, { '∙', "\\bullet" }, { '≀', "\\wr" },
        { '⊓', "\\sqcap" }, { '⊔', "\\sqcup" },
        { '†', "\\dagger" }, { '‡', "\\ddagger" }, { '⨿', "\\amalg" },
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

        foreach (var math in mathPara.Descendants().Where(x => x.LocalName == "oMath"))
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

#pragma warning disable CA1308 // Normalize strings to uppercase
        var localName = element.LocalName.ToLowerInvariant();
#pragma warning restore CA1308 // Normalize strings to uppercase

        if (localName == "t" || element is DocumentFormat.OpenXml.Math.Text || element is DocumentFormat.OpenXml.Wordprocessing.Text)
        {
            return MapText(element.InnerText);
        }

        return localName switch
        {
            "omath" => GetContent(element),
            "omathpara" => GetContent(element),
            "r" => GetContent(element),
            "f" => ParseFraction(element),
            "rad" => string.IsNullOrEmpty(GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("deg", StringComparison.OrdinalIgnoreCase))))
                               ? "\\sqrt{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}"
                               : "\\sqrt[" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("deg", StringComparison.OrdinalIgnoreCase))) + "]{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}",
            "ssup" => "{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}^{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sup", StringComparison.OrdinalIgnoreCase))) + "}",
            "ssub" => "{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}_{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sub", StringComparison.OrdinalIgnoreCase))) + "}",
            "ssubsup" => "{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}_{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sub", StringComparison.OrdinalIgnoreCase))) + "}^{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sup", StringComparison.OrdinalIgnoreCase))) + "}",
            "spre" => "_{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sub", StringComparison.OrdinalIgnoreCase))) + "}^{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sup", StringComparison.OrdinalIgnoreCase))) + "}{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}",
            "d" => ParseDelimiter(element),
            "nary" => ParseNary(element),
            "m" => ParseMatrixStructure(element),
            "func" => ToFunctionName(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("fname", StringComparison.OrdinalIgnoreCase))) + " " + WrapContent(GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase)))),
            "limlow" => ParseLimit(element, true),
            "limupp" => ParseLimit(element, false),
            "acc" => ParseAccent(element),
            "groupchr" => ParseGroupChar(element),
            "bar" => ParseBar(element),
            "box" => "\\boxed{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}",
            "borderbox" => "\\boxed{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}",
            "eqarr" => "\\begin{aligned} " + ParseEquationArray(element) + " \\end{aligned}",
            _ => localName.EndsWith("pr", StringComparison.OrdinalIgnoreCase) ? string.Empty : GetContent(element),
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

    private static string GetValAttribute(OpenXmlElement? element)
    {
        if (element is null)
        {
            return string.Empty;
        }

        // Try to find any attribute with local name "val"
        foreach (var attr in element.GetAttributes())
        {
            if (string.Equals(attr.LocalName, "val", StringComparison.OrdinalIgnoreCase))
            {
                return attr.Value ?? string.Empty;
            }
        }

        return string.Empty;
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

    private static string ParseFraction(OpenXmlElement element)
    {
        var fPr = element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("fpr", StringComparison.OrdinalIgnoreCase));
        var fType = fPr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("ftype", StringComparison.OrdinalIgnoreCase));
        var typeVal = GetValAttribute(fType);

        var num = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("num", StringComparison.OrdinalIgnoreCase)));
        var den = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("den", StringComparison.OrdinalIgnoreCase)));

        if (string.Equals(typeVal, "noBar", StringComparison.OrdinalIgnoreCase))
        {
            return "\\binom{" + num + "}{" + den + "}";
        }

        return "\\frac{" + num + "}{" + den + "}";
    }

    private static string ParseDelimiter(OpenXmlElement element)
    {
        var dPr = element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("dpr", StringComparison.OrdinalIgnoreCase));
        var begChr = GetValAttribute(dPr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("begchr", StringComparison.OrdinalIgnoreCase)));
        var endChr = GetValAttribute(dPr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("endchr", StringComparison.OrdinalIgnoreCase)));

        if (string.IsNullOrEmpty(begChr))
        {
            begChr = "(";
        }

        var content = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase)));

        // Piecewise detection
        if (begChr == "{" && string.IsNullOrEmpty(endChr))
        {
            // If it contains a matrix, use cases
            if (content.Contains("\\begin{matrix}", StringComparison.Ordinal))
            {
                return content.Replace("\\begin{matrix}", "\\begin{cases}", StringComparison.Ordinal).Replace("\\end{matrix}", "\\end{cases}", StringComparison.Ordinal);
            }
        }

        // Map brackets
        var l = MapBracket(begChr);
        var r = string.IsNullOrEmpty(endChr) ? "." : MapBracket(endChr);

        return "\\left" + l + " " + content + " \\right" + r;
    }

    private static string MapBracket(string chr)
    {
        return chr switch
        {
            "{" => "\\{",
            "}" => "\\}",
            "[" => "[",
            "]" => "]",
            "(" => "(",
            ")" => ")",
            "|" => "|",
            "‖" => "\\|",
            "⌊" => "\\lfloor",
            "⌋" => "\\rfloor",
            "⌈" => "\\lceil",
            "⌉" => "\\rceil",
            _ => chr,
        };
    }

    private static string ParseNary(OpenXmlElement element)
    {
        var naryPr = element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("narypr", StringComparison.OrdinalIgnoreCase));
        var chr = GetValAttribute(naryPr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("chr", StringComparison.OrdinalIgnoreCase)));

        if (string.IsNullOrEmpty(chr))
        {
            chr = "∫";
        }

        var op = chr switch
        {
            "∑" => "\\sum",
            "∏" => "\\prod",
            "∫" => "\\int",
            "∬" => "\\iint",
            "∭" => "\\iiint",
            "∮" => "\\oint",
            _ => GetDefaultNaryOperator(chr),
        };

        var sub = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sub", StringComparison.OrdinalIgnoreCase)));
        var sup = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("sup", StringComparison.OrdinalIgnoreCase)));
        var e = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase)));

        var sb = new StringBuilder();
        sb.Append(op);

        if (!string.IsNullOrEmpty(sub))
        {
            sb.Append("_{").Append(sub).Append('}');
        }

        if (!string.IsNullOrEmpty(sup))
        {
            sb.Append("^{").Append(sup).Append('}');
        }

        sb.Append(' ').Append(e);
        return sb.ToString();
    }

    private static string GetDefaultNaryOperator(string chr)
    {
        if (string.IsNullOrEmpty(chr))
        {
            return "\\int";
        }

        if (SymbolMap.TryGetValue(chr[0], out var latex))
        {
            return latex;
        }

        return "\\int";
    }

    private static string ParseAccent(OpenXmlElement element)
    {
        var accPr = element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("accpr", StringComparison.OrdinalIgnoreCase));

        var chr = GetValAttribute(accPr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("chr", StringComparison.OrdinalIgnoreCase)));

        if (string.IsNullOrEmpty(chr))
        {
            chr = "¯";
        }

        var acc = chr switch
        {
            "¯" => "\\bar",
            "⃗" => "\\vec",
            "̇" => "\\dot",
            "̈" => "\\ddot",
            "̂" => "\\hat",
            "̃" => "\\tilde",
            "̌" => "\\check",
            "̆" => "\\breve",
            "́" => "\\acute",
            "̀" => "\\grave",
            _ => "\\bar",
        };

        return acc + "{" + GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase))) + "}";
    }

    private static string ParseLimit(OpenXmlElement limitElement, bool isLower)
    {
        var e = limitElement.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase));
        var lim = limitElement.ChildElements.FirstOrDefault(x => x.LocalName.Equals("lim", StringComparison.OrdinalIgnoreCase));

        var eContent = GetContent(e).Trim();
        var limContent = GetContent(lim).Trim();

        if (string.Equals(eContent, "lim", StringComparison.OrdinalIgnoreCase) || eContent.Contains("\\lim", StringComparison.Ordinal))
        {
            return (isLower ? "\\lim_{" : "\\overline{lim}_{") + limContent + "}";
        }

        return "{" + eContent + "}" + (isLower ? "_{" : "^{") + limContent + "}";
    }

    private static string ParseMatrixStructure(OpenXmlElement matrix)
    {
        var sb = new StringBuilder();
        sb.Append("\\begin{matrix} ");
        var rows = matrix.ChildElements.Where(x => x.LocalName.Equals("mr", StringComparison.OrdinalIgnoreCase)).ToList();
        for (var i = 0; i < rows.Count; i++)
        {
            // Rows can contain 'e' elements directly or 'arg' elements
            var cells = rows[i].ChildElements.Where(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase) || x.LocalName.Equals("arg", StringComparison.OrdinalIgnoreCase)).ToList();
            if (cells.Count == 0)
            {
                // Backup: check all children that are not properties
                cells = rows[i].ChildElements.Where(x => !x.LocalName.EndsWith("pr", StringComparison.OrdinalIgnoreCase)).ToList();
            }

            for (var j = 0; j < cells.Count; j++)
            {
                sb.Append(GetContent(cells[j]));
                if (j < cells.Count - 1)
                {
                    sb.Append(" & ");
                }
            }

            if (i < rows.Count - 1)
            {
                sb.Append(" \\\\ ");
            }
        }

        sb.Append(" \\end{matrix}");
        return sb.ToString();
    }

    private static string ParseGroupChar(OpenXmlElement element)
    {
        var pr = element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("groupchrpr", StringComparison.OrdinalIgnoreCase));
        var chr = GetValAttribute(pr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("chr", StringComparison.OrdinalIgnoreCase)));
        var pos = GetValAttribute(pr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("pos", StringComparison.OrdinalIgnoreCase)));

        if (string.IsNullOrEmpty(chr))
        {
            chr = "→";
        }

        var latexChr = MapText(chr).Trim();
        var baseExpr = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase)));

        if (pos == "bot")
        {
            return "\\underset{" + latexChr + "}{" + baseExpr + "}";
        }

        return "\\overset{" + latexChr + "}{" + baseExpr + "}";
    }

    private static string ParseBar(OpenXmlElement element)
    {
        var pr = element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("barpr", StringComparison.OrdinalIgnoreCase));
        var pos = GetValAttribute(pr?.ChildElements.FirstOrDefault(x => x.LocalName.Equals("pos", StringComparison.OrdinalIgnoreCase)));
        var e = GetContent(element.ChildElements.FirstOrDefault(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase)));

        if (pos == "bot")
        {
            return "\\underline{" + e + "}";
        }

        return "\\overline{" + e + "}";
    }

    private static string ParseEquationArray(OpenXmlElement element)
    {
        var sb = new StringBuilder();
        var rows = element.ChildElements.Where(x => x.LocalName.Equals("e", StringComparison.OrdinalIgnoreCase)).ToList();
        for (var i = 0; i < rows.Count; i++)
        {
            sb.Append(ToLatex(rows[i]));
            if (i < rows.Count - 1)
            {
                sb.Append(" \\\\ ");
            }
        }

        return sb.ToString();
    }
}
