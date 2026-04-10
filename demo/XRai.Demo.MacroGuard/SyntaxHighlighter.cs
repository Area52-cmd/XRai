using System.Text.RegularExpressions;
using System.Windows.Documents;
using System.Windows.Media;
using Color = System.Windows.Media.Color;

namespace XRai.Demo.MacroGuard;

public static class SyntaxHighlighter
{
    private static readonly string[] VbaKeywords =
    {
        "Sub", "End Sub", "Function", "End Function", "Property", "End Property",
        "If", "Then", "Else", "ElseIf", "End If",
        "For", "To", "Step", "Next", "Each", "In",
        "Do", "Loop", "While", "Wend", "Until",
        "With", "End With",
        "Select", "Case", "End Select",
        "Dim", "Set", "Let", "As", "New", "Nothing",
        "Public", "Private", "Static", "Const",
        "ByVal", "ByRef", "Optional", "ParamArray",
        "On", "Error", "GoTo", "Resume",
        "Exit", "Return",
        "True", "False", "Not", "And", "Or", "Xor",
        "Is", "Like", "Mod",
        "Call", "ReDim", "Preserve",
        "Option", "Explicit", "Compare", "Base",
        "Type", "Enum", "End Type", "End Enum",
        "Boolean", "Integer", "Long", "Double", "Single", "String", "Variant", "Object", "Date", "Currency", "Byte",
        "Me", "Debug", "Print",
        "MsgBox", "InputBox"
    };

    private static readonly System.Windows.Media.Brush KeywordBrush = new SolidColorBrush(Color.FromRgb(86, 156, 214));   // Blue
    private static readonly System.Windows.Media.Brush CommentBrush = new SolidColorBrush(Color.FromRgb(106, 153, 85));   // Green
    private static readonly System.Windows.Media.Brush StringBrush = new SolidColorBrush(Color.FromRgb(206, 145, 120));    // Red/Orange
    private static readonly System.Windows.Media.Brush NumberBrush = new SolidColorBrush(Color.FromRgb(181, 206, 168));    // Light green
    private static readonly System.Windows.Media.Brush DefaultBrush = new SolidColorBrush(Color.FromRgb(212, 212, 212));   // Light gray

    private static readonly Regex StringPattern = new(@"""[^""]*""", RegexOptions.Compiled);
    private static readonly Regex NumberPattern = new(@"\b\d+\.?\d*\b", RegexOptions.Compiled);
    private static readonly Regex CommentPattern = new(@"'.*$", RegexOptions.Compiled | RegexOptions.Multiline);

    static SyntaxHighlighter()
    {
        KeywordBrush.Freeze();
        CommentBrush.Freeze();
        StringBrush.Freeze();
        NumberBrush.Freeze();
        DefaultBrush.Freeze();
    }

    public static FlowDocument Highlight(string code)
    {
        var doc = new FlowDocument
        {
            Background = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
            Foreground = DefaultBrush,
            FontFamily = new System.Windows.Media.FontFamily("Cascadia Code, Consolas, Courier New"),
            FontSize = 12,
            PagePadding = new System.Windows.Thickness(8)
        };

        var paragraph = new Paragraph();

        foreach (var line in code.Split('\n'))
        {
            HighlightLine(paragraph, line);
            paragraph.Inlines.Add(new LineBreak());
        }

        doc.Blocks.Add(paragraph);
        return doc;
    }

    private static void HighlightLine(Paragraph paragraph, string line)
    {
        // Check for comment first
        var commentMatch = CommentPattern.Match(line);
        string codePart = commentMatch.Success ? line[..commentMatch.Index] : line;
        string? commentPart = commentMatch.Success ? commentMatch.Value : null;

        // Process the code part
        if (!string.IsNullOrEmpty(codePart))
        {
            HighlightCodeSegment(paragraph, codePart);
        }

        // Append comment in green
        if (commentPart != null)
        {
            paragraph.Inlines.Add(new Run(commentPart) { Foreground = CommentBrush });
        }
    }

    private static void HighlightCodeSegment(Paragraph paragraph, string code)
    {
        // Find all string literals first to avoid keyword matching inside strings
        var stringRanges = new List<(int Start, int End)>();
        foreach (Match m in StringPattern.Matches(code))
        {
            stringRanges.Add((m.Index, m.Index + m.Length));
        }

        int pos = 0;
        while (pos < code.Length)
        {
            // Check if we're inside a string
            var inString = stringRanges.FirstOrDefault(r => pos >= r.Start && pos < r.End);
            if (inString != default)
            {
                paragraph.Inlines.Add(new Run(code[inString.Start..inString.End]) { Foreground = StringBrush });
                pos = inString.End;
                continue;
            }

            // Try to match a keyword at this position
            bool matched = false;
            if (char.IsLetter(code[pos]))
            {
                foreach (var kw in VbaKeywords)
                {
                    if (pos + kw.Length > code.Length) continue;
                    if (!code.AsSpan(pos, kw.Length).Equals(kw.AsSpan(), StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Check word boundary
                    bool startBoundary = pos == 0 || !char.IsLetterOrDigit(code[pos - 1]);
                    bool endBoundary = pos + kw.Length >= code.Length || !char.IsLetterOrDigit(code[pos + kw.Length]);
                    // For multi-word keywords like "End Sub", we allow the space
                    if (kw.Contains(' '))
                        endBoundary = pos + kw.Length >= code.Length || !char.IsLetterOrDigit(code[pos + kw.Length]);

                    if (startBoundary && endBoundary)
                    {
                        paragraph.Inlines.Add(new Run(code.Substring(pos, kw.Length)) { Foreground = KeywordBrush });
                        pos += kw.Length;
                        matched = true;
                        break;
                    }
                }
            }

            if (matched) continue;

            // Try number
            if (char.IsDigit(code[pos]))
            {
                int start = pos;
                while (pos < code.Length && (char.IsDigit(code[pos]) || code[pos] == '.'))
                    pos++;
                paragraph.Inlines.Add(new Run(code[start..pos]) { Foreground = NumberBrush });
                continue;
            }

            // Default character
            int defStart = pos;
            while (pos < code.Length && !char.IsLetter(code[pos]) && !char.IsDigit(code[pos]) &&
                   !stringRanges.Any(r => pos >= r.Start && pos < r.End))
            {
                pos++;
                // Break if next char could start a keyword/number/string
                if (pos < code.Length && (char.IsLetter(code[pos]) || char.IsDigit(code[pos]) || code[pos] == '"'))
                    break;
            }
            if (pos > defStart)
                paragraph.Inlines.Add(new Run(code[defStart..pos]) { Foreground = DefaultBrush });
            else if (!matched)
            {
                // Single unmatched letter character — advance to avoid infinite loop
                int wordStart = pos;
                while (pos < code.Length && char.IsLetterOrDigit(code[pos]))
                    pos++;
                paragraph.Inlines.Add(new Run(code[wordStart..pos]) { Foreground = DefaultBrush });
            }
        }
    }
}
