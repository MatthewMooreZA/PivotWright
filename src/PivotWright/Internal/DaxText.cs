namespace PivotWright.Internal;

/// <summary>
/// A shallow check of DAX text: brackets, parentheses, braces and quotes must be paired. It does not parse DAX; it only catches text
/// that Excel would reject as malformed, so the mistake shows up where the formula is written rather than in the PivotTable.
/// </summary>
internal static class DaxText
{
    /// <summary>The first problem found, with its 1-based position, or null when the text is well-formed.</summary>
    public static string? Problem(string text)
    {
        var open = new Stack<(char Symbol, int Position)>();
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            char next = i + 1 < text.Length ? text[i + 1] : '\0';
            switch (c)
            {
                case '/' when next == '/':
                case '-' when next == '-':
                    while (i < text.Length && text[i] != '\n') i++;
                    break;
                case '/' when next == '*':
                {
                    int end = text.IndexOf("*/", i + 2, StringComparison.Ordinal);
                    if (end < 0) return $"the comment that starts at position {i + 1} is not closed";
                    i = end + 1;
                    break;
                }
                case '"':
                    i = SkipQuoted(text, i, '"', "text in double quotes", out string? stringProblem);
                    if (stringProblem is not null) return stringProblem;
                    break;
                case '\'':
                    i = SkipQuoted(text, i, '\'', "the table name in single quotes", out string? tableProblem);
                    if (tableProblem is not null) return tableProblem;
                    break;
                case '[':
                {
                    // A column or measure name: ends at the first ] that is not doubled.
                    int j = i + 1;
                    while (j < text.Length && !(text[j] == ']' && (j + 1 >= text.Length || text[j + 1] != ']'))) j += text[j] == ']' ? 2 : 1;
                    if (j >= text.Length) return $"the bracket at position {i + 1} is not closed";
                    i = j;
                    break;
                }
                case '(':
                case '{':
                    open.Push((c, i + 1));
                    break;
                case ')':
                case '}':
                {
                    char expected = c == ')' ? '(' : '{';
                    if (open.Count == 0 || open.Peek().Symbol != expected) return $"there is a {c} at position {i + 1} without a matching {expected}";
                    open.Pop();
                    break;
                }
                case ']':
                    return $"there is a ] at position {i + 1} without a matching [";
            }
        }
        if (open.Count > 0)
        {
            var (symbol, position) = open.Peek();
            return $"the {symbol} at position {position} is not closed";
        }
        return null;
    }

    private static int SkipQuoted(string text, int start, char quote, string what, out string? problem)
    {
        problem = null;
        for (int j = start + 1; j < text.Length; j++)
        {
            if (text[j] != quote) continue;
            if (j + 1 < text.Length && text[j + 1] == quote) { j++; continue; }
            return j;
        }
        problem = $"{what} that starts at position {start + 1} is not closed";
        return text.Length;
    }

    /// <summary>Throws an ArgumentException naming the problem when the text is malformed.</summary>
    public static void Check(string expression, string paramName)
    {
        if (Problem(expression) is { } problem)
            throw new ArgumentException($"The DAX expression '{expression}' is malformed: {problem}.", paramName);
    }
}
