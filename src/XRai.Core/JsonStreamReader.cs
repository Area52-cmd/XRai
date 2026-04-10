using System.Text;

namespace XRai.Core;

/// <summary>
/// Character-mode JSON document reader that handles pretty-printed multi-line JSON.
/// Tracks brace/bracket depth outside string literals, respects escape sequences,
/// and yields complete JSON documents as they close at the top level.
///
/// Used by both Repl (direct stdin dispatch) and DaemonClient (forwarding stdin
/// to the daemon) so the parsing behavior is identical regardless of whether the
/// daemon is running. Previously DaemonClient used ReadLine which broke multi-line
/// JSON — fixed in Round 9 by making it use this helper.
/// </summary>
public static class JsonStreamReader
{
    /// <summary>
    /// Read complete JSON documents from a TextReader. Each call to the returned
    /// enumerable yields one complete top-level JSON object or array as a string.
    /// Returns when the reader hits EOF. Skips whitespace between documents.
    /// </summary>
    public static IEnumerable<string> ReadDocuments(TextReader reader)
    {
        var buf = new StringBuilder();
        int depth = 0;
        bool inString = false;
        bool escape = false;

        while (true)
        {
            int ch;
            try { ch = reader.Read(); }
            catch { yield break; }

            if (ch == -1)
            {
                // EOF — emit any trailing partial document if it's actually complete
                if (buf.Length > 0 && depth == 0)
                {
                    var final = buf.ToString().Trim();
                    if (final.Length > 0) yield return final;
                }
                yield break;
            }

            char c = (char)ch;

            // String state tracking
            if (escape)
            {
                escape = false;
                buf.Append(c);
                continue;
            }
            if (c == '\\' && inString)
            {
                escape = true;
                buf.Append(c);
                continue;
            }
            if (c == '"')
            {
                inString = !inString;
                buf.Append(c);
                continue;
            }

            if (inString)
            {
                buf.Append(c);
                continue;
            }

            // Structural characters outside strings
            if (c == '{' || c == '[')
            {
                depth++;
                buf.Append(c);
                continue;
            }
            if (c == '}' || c == ']')
            {
                depth--;
                buf.Append(c);

                if (depth == 0)
                {
                    var payload = buf.ToString().Trim();
                    buf.Clear();
                    if (payload.Length > 0)
                        yield return payload;
                }
                continue;
            }

            // Whitespace
            if (char.IsWhiteSpace(c))
            {
                if (depth > 0 || buf.Length > 0)
                    buf.Append(c);
                continue;
            }

            buf.Append(c);
        }
    }
}
