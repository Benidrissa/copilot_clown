using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using CopilotClown.Functions;

namespace CopilotClown.UI;

/// <summary>
/// Stores formulas before Convert to Values so they can be restored (undo).
/// Only the last conversion is kept.
/// </summary>
internal static class ConvertUndoState
{
    // Key: "SheetName!A1" → formula string
    internal static Dictionary<string, string> SavedFormulas = new Dictionary<string, string>();

    internal static void Clear() => SavedFormulas.Clear();

    internal static bool HasUndo => SavedFormulas.Count > 0;
}

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string ribbonID)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab idMso='TabHome'>
        <group id='CopilotClownGroup' label='AI Assistant'>
          <button id='btnSettings'
                  label='AI Settings'
                  size='large'
                  imageMso='PropertySheet'
                  onAction='OnSettingsClick'
                  screentip='Configure AI Settings'
                  supertip='Set API keys, choose AI model, and manage response cache for the =USEAI() function.' />
          <button id='btnConvertValues'
                  label='Convert to Values'
                  size='large'
                  imageMso='PasteValues'
                  onAction='OnConvertValuesClick'
                  screentip='Convert Selection to Values'
                  supertip='Replace formulas in selected cells with their computed text values, so the workbook can be shared. Use Undo Convert to restore. Shortcut: Ctrl+Shift+V' />
          <button id='btnUndoConvert'
                  label='Undo Convert'
                  size='large'
                  imageMso='Undo'
                  onAction='OnUndoConvertClick'
                  screentip='Undo Convert to Values'
                  supertip='Restore the USEAI formulas that were replaced by the last Convert to Values operation.' />
          <separator id='sepRefresh' />
          <button id='btnRefreshAll'
                  label='Refresh All'
                  size='normal'
                  imageMso='RefreshAll'
                  onAction='OnRefreshAllClick'
                  screentip='Refresh All USEAI Cells'
                  supertip='Clear the AI response cache and recalculate all USEAI formulas. Respects rate limits.' />
          <button id='btnRefreshSelected'
                  label='Refresh Selected'
                  size='normal'
                  imageMso='Refresh'
                  onAction='OnRefreshSelectedClick'
                  screentip='Refresh Selected USEAI Cells'
                  supertip='Clear cached results for selected cells and recalculate them. Respects rate limits.' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public void OnSettingsClick(IRibbonControl control)
    {
        var form = new SettingsForm();
        form.ShowDialog();
    }

    public void OnConvertValuesClick(IRibbonControl control)
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic selection = app.Selection;

            // Save formulas for undo before converting
            ConvertUndoState.Clear();
            foreach (dynamic cell in selection.Cells)
            {
                try
                {
                    if (cell.HasFormula)
                    {
                        string sheetName = cell.Worksheet.Name;
                        string address = cell.Address;
                        string formula = cell.Formula;
                        ConvertUndoState.SavedFormulas[$"{sheetName}!{address}"] = formula;
                    }
                }
                catch { }
            }

            selection.Value2 = selection.Value2;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not convert: {ex.Message}",
                "Convert to Values",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public void OnUndoConvertClick(IRibbonControl control)
    {
        if (!ConvertUndoState.HasUndo)
        {
            MessageBox.Show(
                "Nothing to undo. Convert to Values has not been used yet.",
                "Undo Convert",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic workbook = app.ActiveWorkbook;
            int restored = 0;

            foreach (var kvp in ConvertUndoState.SavedFormulas)
            {
                try
                {
                    // Key format: "SheetName!$A$1"
                    int sep = kvp.Key.IndexOf('!');
                    string sheetName = kvp.Key.Substring(0, sep);
                    string address = kvp.Key.Substring(sep + 1);

                    dynamic sheet = workbook.Worksheets[sheetName];
                    dynamic cell = sheet.Range[address];
                    cell.Formula = kvp.Value;
                    restored++;
                }
                catch { }
            }

            ConvertUndoState.Clear();

            MessageBox.Show(
                $"Restored {restored} formula(s).",
                "Undo Convert",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Undo failed: {ex.Message}",
                "Undo Convert",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public void OnRefreshAllClick(IRibbonControl control)
    {
        var result = MessageBox.Show(
            "This will clear all cached AI responses and recalculate every USEAI formula in the workbook.\n\n" +
            "This may trigger many API calls (rate limits apply). Continue?",
            "Refresh All",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes) return;

        try
        {
            UseAiFunction.RefreshAll();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Refresh failed: {ex.Message}",
                "Refresh All",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    public void OnRefreshSelectedClick(IRibbonControl control)
    {
        try
        {
            var (refreshed, _) = UseAiFunction.RefreshSelected();
            if (refreshed == 0)
            {
                MessageBox.Show(
                    "No USEAI formulas found in the current selection.",
                    "Refresh Selected",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Refresh failed: {ex.Message}",
                "Refresh Selected",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}

/// <summary>
/// Macro commands for AI Assistant — work even if ribbon doesn't load.
/// Run from Excel: Alt+F8 > ShowAISettings / ConvertToValues
/// </summary>
public static class Commands
{
    [ExcelCommand(MenuName = "AI Assistant", MenuText = "AI Settings")]
    public static void ShowAISettings()
    {
        var form = new SettingsForm();
        form.ShowDialog();
    }

    [ExcelCommand(
        MenuName = "AI Assistant",
        MenuText = "Convert to Values",
        ShortCut = "^+V")]
    public static void ConvertToValues()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic selection = app.Selection;

            // Save formulas for undo
            ConvertUndoState.Clear();
            foreach (dynamic cell in selection.Cells)
            {
                try
                {
                    if (cell.HasFormula)
                    {
                        string sheetName = cell.Worksheet.Name;
                        string address = cell.Address;
                        string formula = cell.Formula;
                        ConvertUndoState.SavedFormulas[$"{sheetName}!{address}"] = formula;
                    }
                }
                catch { }
            }

            selection.Value2 = selection.Value2;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not convert: {ex.Message}",
                "Convert to Values",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}
