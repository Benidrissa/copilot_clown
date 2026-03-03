using System.Text;

namespace CopilotClown.Services;

public static class PromptBuilder
{
    /// <summary>
    /// Build a prompt string from the variadic arguments passed to =USEAI().
    /// Uses StringBuilder to minimize allocations.
    /// </summary>
    public static string Build(object[] args)
    {
        if (args.Length == 0) return "";

        var sb = new StringBuilder(256);

        foreach (var arg in args)
        {
            switch (arg)
            {
                case string s:
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(s);
                    break;
                case double d:
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(d.ToString("G"));
                    break;
                case bool b:
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(b);
                    break;
                case object[,] range:
                    if (sb.Length > 0) sb.Append(' ');
                    FlattenRange(range, sb);
                    break;
                case ExcelDna.Integration.ExcelMissing:
                case ExcelDna.Integration.ExcelEmpty:
                    break;
                default:
                    if (arg != null)
                    {
                        if (sb.Length > 0) sb.Append(' ');
                        sb.Append(arg);
                    }
                    break;
            }
        }

        return sb.ToString().Trim();
    }

    private static void FlattenRange(object[,] range, StringBuilder sb)
    {
        var rows = range.GetLength(0);
        var cols = range.GetLength(1);
        if (rows == 0 || cols == 0) return;

        bool firstLine = true;
        for (int r = 0; r < rows; r++)
        {
            bool hasContent = false;
            var lineStart = sb.Length;

            if (!firstLine) sb.Append('\n');

            for (int c = 0; c < cols; c++)
            {
                var val = range[r, c];
                if (val is ExcelDna.Integration.ExcelEmpty || val == null) continue;
                if (hasContent) sb.Append(", ");
                sb.Append(val);
                hasContent = true;
            }

            if (!hasContent)
                sb.Length = lineStart; // Remove the newline we added
            else
                firstLine = false;
        }
    }
}
