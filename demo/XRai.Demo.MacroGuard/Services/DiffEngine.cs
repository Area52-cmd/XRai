namespace XRai.Demo.MacroGuard.Services;

/// <summary>
/// Simple line-by-line diff using longest common subsequence.
/// </summary>
public static class DiffEngine
{
    public enum DiffKind { Unchanged, Added, Removed }

    public record DiffLine(DiffKind Kind, string Text);

    public static List<DiffLine> ComputeDiff(string oldText, string newText)
    {
        var oldLines = (oldText ?? "").Split('\n');
        var newLines = (newText ?? "").Split('\n');
        return ComputeDiff(oldLines, newLines);
    }

    public static List<DiffLine> ComputeDiff(string[] oldLines, string[] newLines)
    {
        int m = oldLines.Length, n = newLines.Length;

        // Build LCS length table
        var dp = new int[m + 1, n + 1];
        for (int i = 1; i <= m; i++)
            for (int j = 1; j <= n; j++)
                dp[i, j] = oldLines[i - 1] == newLines[j - 1]
                    ? dp[i - 1, j - 1] + 1
                    : Math.Max(dp[i - 1, j], dp[i, j - 1]);

        // Backtrack to produce diff
        var result = new List<DiffLine>();
        int ii = m, jj = n;
        while (ii > 0 || jj > 0)
        {
            if (ii > 0 && jj > 0 && oldLines[ii - 1] == newLines[jj - 1])
            {
                result.Add(new DiffLine(DiffKind.Unchanged, oldLines[ii - 1]));
                ii--; jj--;
            }
            else if (jj > 0 && (ii == 0 || dp[ii, jj - 1] >= dp[ii - 1, jj]))
            {
                result.Add(new DiffLine(DiffKind.Added, newLines[jj - 1]));
                jj--;
            }
            else
            {
                result.Add(new DiffLine(DiffKind.Removed, oldLines[ii - 1]));
                ii--;
            }
        }

        result.Reverse();
        return result;
    }
}
