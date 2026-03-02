using System.Collections.Generic;

namespace CopilotClown.Services;

public static class PromptBuilder
{
    /// <summary>
    /// Build a prompt string from the variadic arguments passed to =USEAI().
    /// Arguments alternate between prompt text (string) and context values.
    /// Context from Excel ranges arrives as object[,] (2D array).
    /// </summary>
    public static string Build(object[] args)
    {
        if (args.Length == 0) return "";

        var parts = new List<string>();

        foreach (var arg in args)
        {
            switch (arg)
            {
                case string s:
                    parts.Add(s);
                    break;
                case double d:
                    parts.Add(d.ToString("G"));
                    break;
                case bool b:
                    parts.Add(b.ToString());
                    break;
                case object[,] range:
                    parts.Add(FlattenRange(range));
                    break;
                case ExcelDna.Integration.ExcelMissing:
                    // Skip missing optional arguments
                    break;
                case ExcelDna.Integration.ExcelEmpty:
                    // Skip empty cells
                    break;
                default:
                    if (arg != null)
                        parts.Add(arg.ToString() ?? "");
                    break;
            }
        }

        return string.Join(" ", parts).Trim();
    }

    private static string FlattenRange(object[,] range)
    {
        var rows = range.GetLength(0);
        var cols = range.GetLength(1);

        if (rows == 0 || cols == 0) return "";

        var lines = new List<string>();

        for (int r = 0; r < rows; r++)
        {
            var cells = new List<string>();
            for (int c = 0; c < cols; c++)
            {
                var val = range[r, c];
                if (val is not ExcelDna.Integration.ExcelEmpty and not null)
                    cells.Add(val.ToString() ?? "");
            }
            if (cells.Count > 0)
            {
                lines.Add(cols == 1 ? cells[0] : string.Join(", ", cells));
            }
        }

        return string.Join("\n", lines);
    }
}
