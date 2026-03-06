using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

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
                  supertip='Replace formulas in selected cells with their computed text values, so the workbook can be shared.' />
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
}

/// <summary>
/// Macro command to open settings — works even if ribbon doesn't load.
/// Run from Excel: Alt+F8 > ShowAISettings
/// </summary>
public static class Commands
{
    [ExcelCommand(MenuName = "AI Assistant", MenuText = "AI Settings")]
    public static void ShowAISettings()
    {
        var form = new SettingsForm();
        form.ShowDialog();
    }
}
