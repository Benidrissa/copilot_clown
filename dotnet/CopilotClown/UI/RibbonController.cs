using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using CopilotClown.Functions;

namespace CopilotClown.UI;

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
                  supertip='Replace formulas in selected cells with their computed text values, so the workbook can be shared. Shortcut: Ctrl+Shift+V' />
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
