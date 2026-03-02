/**
 * Builds a prompt string from interleaved prompt parts and context values.
 *
 * The =USEAI() function receives alternating string parts and resolved cell/range values.
 * This service stitches them together into a single prompt string.
 */
export class PromptBuilder {
  /**
   * Build a prompt from the raw arguments passed to =USEAI().
   *
   * Arguments alternate between prompt text (string) and context values.
   * Context values can be:
   * - A single string (from a single cell)
   * - A 2D array of strings (from a range)
   *
   * Example:
   *   buildPrompt(["Classify", [["cat","dog","bird"]], "into categories", [["animal","food"]]])
   *   → "Classify cat, dog, bird into categories animal, food"
   */
  buildPrompt(args: unknown[]): string {
    if (args.length === 0) return "";

    const parts: string[] = [];

    for (const arg of args) {
      if (typeof arg === "string") {
        parts.push(arg);
      } else if (typeof arg === "number" || typeof arg === "boolean") {
        parts.push(String(arg));
      } else if (Array.isArray(arg)) {
        parts.push(this.flattenRange(arg));
      } else if (arg !== null && arg !== undefined) {
        parts.push(String(arg));
      }
    }

    return parts.join(" ").trim();
  }

  /**
   * Flatten a 2D array (Excel range) into a readable string.
   * Single-column ranges use newline separation.
   * Multi-column ranges use comma separation within rows, newlines between rows.
   */
  private flattenRange(range: unknown[][]): string {
    if (!Array.isArray(range) || range.length === 0) return "";

    // Check if it's truly 2D
    if (!Array.isArray(range[0])) {
      // 1D array — just join
      return (range as unknown[]).map(String).join(", ");
    }

    const rows = range as unknown[][];
    const isSingleColumn = rows.every((row) => Array.isArray(row) && row.length === 1);

    if (isSingleColumn) {
      return rows.map((row) => String(row[0])).join("\n");
    }

    return rows.map((row) => (Array.isArray(row) ? row.map(String).join(", ") : String(row))).join("\n");
  }
}
