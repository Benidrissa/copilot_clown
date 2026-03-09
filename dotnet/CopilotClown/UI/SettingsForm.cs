using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using CopilotClown.Functions;
using CopilotClown.Models;
using CopilotClown.Services;

namespace CopilotClown.UI;

public class SettingsForm : Form
{
    private readonly SettingsService _settings;
    private readonly CacheService _cache;

    // API Keys tab
    private TextBox _txtClaudeKey;
    private TextBox _txtOpenAIKey;
    private Label _lblClaudeStatus;
    private Label _lblOpenAIStatus;

    // Model tab
    private RadioButton _rbClaude;
    private RadioButton _rbOpenAI;
    private ComboBox _cboModel;
    private Label _lblModelInfo;

    // Cache tab
    private Label _lblCacheStats;
    private ComboBox _cboCacheTtl;
    private CheckBox _chkCacheEnabled;

    // Tools tab
    private Label _lblToolsStatus;

    public SettingsForm()
    {
        _settings = UseAiFunction.SettingsInstance;
        _cache = UseAiFunction.CacheInstance;
        InitializeComponents();
        LoadCurrentSettings();
    }

    private void InitializeComponents()
    {
        Text = "Copilot Clown -- AI Settings";
        Size = new Size(460, 480);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;

        var tabs = new TabControl { Dock = DockStyle.Fill };

        tabs.TabPages.Add(CreateApiKeysTab());
        tabs.TabPages.Add(CreateModelTab());
        tabs.TabPages.Add(CreateCacheTab());
        tabs.TabPages.Add(CreateToolsTab());
        tabs.TabPages.Add(CreateAboutTab());

        Controls.Add(tabs);
    }

    // ── API Keys Tab ────────────────────────────────────────────────

    private TabPage CreateApiKeysTab()
    {
        var page = new TabPage("API Keys") { Padding = new Padding(12) };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 8,
            AutoSize = true,
        };

        // Claude section
        layout.Controls.Add(new Label { Text = "Anthropic (Claude) API Key:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _txtClaudeKey = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
        layout.Controls.Add(_txtClaudeKey);

        var claudeBtnPanel = new FlowLayoutPanel { AutoSize = true };
        var btnSaveClaude = new Button { Text = "Save", Width = 70 };
        btnSaveClaude.Click += (s, e) => SaveApiKey(ProviderName.Anthropic, _txtClaudeKey.Text);
        var btnTestClaude = new Button { Text = "Test", Width = 70 };
        btnTestClaude.Click += async (s, e) => await TestApiKey(ProviderName.Anthropic, _txtClaudeKey.Text, _lblClaudeStatus);
        claudeBtnPanel.Controls.AddRange(new Control[] { btnSaveClaude, btnTestClaude });
        layout.Controls.Add(claudeBtnPanel);
        _lblClaudeStatus = new Label { AutoSize = true, ForeColor = Color.Gray, Text = "" };
        layout.Controls.Add(_lblClaudeStatus);

        // Spacer
        layout.Controls.Add(new Label { Height = 16 });

        // OpenAI section
        layout.Controls.Add(new Label { Text = "OpenAI API Key:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _txtOpenAIKey = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
        layout.Controls.Add(_txtOpenAIKey);

        var oaiBtnPanel = new FlowLayoutPanel { AutoSize = true };
        var btnSaveOAI = new Button { Text = "Save", Width = 70 };
        btnSaveOAI.Click += (s, e) => SaveApiKey(ProviderName.OpenAI, _txtOpenAIKey.Text);
        var btnTestOAI = new Button { Text = "Test", Width = 70 };
        btnTestOAI.Click += async (s, e) => await TestApiKey(ProviderName.OpenAI, _txtOpenAIKey.Text, _lblOpenAIStatus);
        oaiBtnPanel.Controls.AddRange(new Control[] { btnSaveOAI, btnTestOAI });
        layout.Controls.Add(oaiBtnPanel);
        _lblOpenAIStatus = new Label { AutoSize = true, ForeColor = Color.Gray, Text = "" };
        layout.Controls.Add(_lblOpenAIStatus);

        page.Controls.Add(layout);
        return page;
    }

    // ── Model Tab ───────────────────────────────────────────────────

    private TabPage CreateModelTab()
    {
        var page = new TabPage("Model") { Padding = new Padding(12) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6 };

        layout.Controls.Add(new Label { Text = "Provider:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });

        var providerPanel = new FlowLayoutPanel { AutoSize = true };
        _rbClaude = new RadioButton { Text = "Claude (Anthropic)", AutoSize = true };
        _rbOpenAI = new RadioButton { Text = "OpenAI", AutoSize = true };
        _rbClaude.CheckedChanged += (s, e) => RefreshModelList();
        _rbOpenAI.CheckedChanged += (s, e) => RefreshModelList();
        providerPanel.Controls.AddRange(new Control[] { _rbClaude, _rbOpenAI });
        layout.Controls.Add(providerPanel);

        layout.Controls.Add(new Label { Text = "Model:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _cboModel = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        _cboModel.SelectedIndexChanged += (s, e) => RefreshModelInfo();
        layout.Controls.Add(_cboModel);

        _lblModelInfo = new Label { AutoSize = true, ForeColor = Color.Gray };
        layout.Controls.Add(_lblModelInfo);

        var btnSaveModel = new Button { Text = "Save Selection", Width = 120 };
        btnSaveModel.Click += (s, e) => SaveModelSelection();
        layout.Controls.Add(btnSaveModel);

        page.Controls.Add(layout);
        return page;
    }

    // ── Cache Tab ───────────────────────────────────────────────────

    private TabPage CreateCacheTab()
    {
        var page = new TabPage("Cache") { Padding = new Padding(12) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6 };

        _chkCacheEnabled = new CheckBox { Text = "Enable response caching", AutoSize = true };
        layout.Controls.Add(_chkCacheEnabled);

        _lblCacheStats = new Label { AutoSize = true, Font = new Font("Consolas", 9) };
        layout.Controls.Add(_lblCacheStats);

        layout.Controls.Add(new Label { Text = "Cache TTL:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _cboCacheTtl = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        _cboCacheTtl.Items.AddRange(new object[] { "1 hour", "6 hours", "24 hours", "7 days", "30 days" });
        layout.Controls.Add(_cboCacheTtl);

        var btnPanel = new FlowLayoutPanel { AutoSize = true };
        var btnClear = new Button { Text = "Clear API Cache", Width = 120, ForeColor = Color.DarkRed };
        btnClear.Click += (s, e) => { _cache.Clear(); RefreshCacheStats(); };
        var btnClearFiles = new Button { Text = "Clear File Cache", Width = 120, ForeColor = Color.DarkRed };
        btnClearFiles.Click += (s, e) =>
        {
            UseAiFunction.FileCacheInstance.Clear();
            UseAiFunction.UploadCacheInstance.Clear();
            RefreshCacheStats();
            MessageBox.Show("File content and upload caches cleared.", "Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
        var btnSaveCache = new Button { Text = "Save Settings", Width = 120 };
        btnSaveCache.Click += (s, e) => SaveCacheSettings();
        btnPanel.Controls.AddRange(new Control[] { btnClear, btnClearFiles, btnSaveCache });
        layout.Controls.Add(btnPanel);

        page.Controls.Add(layout);
        return page;
    }

    // ── Tools Tab ─────────────────────────────────────────────────

    private TabPage CreateToolsTab()
    {
        var page = new TabPage("Tools") { Padding = new Padding(12) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };

        layout.Controls.Add(new Label
        {
            Text = "Share Workbook",
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold),
        });

        layout.Controls.Add(new Label
        {
            Text = "Convert USEAI formulas to plain values so the workbook\ncan be shared with users who don't have the add-in installed.",
            AutoSize = true,
        });

        var btnPanel = new FlowLayoutPanel { AutoSize = true };

        var btnConvertSelected = new Button { Text = "Convert Selected Cells", Width = 160 };
        btnConvertSelected.Click += (s, e) => ConvertSelectedCells();

        var btnConvertAll = new Button { Text = "Convert All USEAI Cells", Width = 170 };
        btnConvertAll.Click += (s, e) => ConvertAllUseAiCells();

        btnPanel.Controls.AddRange(new Control[] { btnConvertSelected, btnConvertAll });
        layout.Controls.Add(btnPanel);

        _lblToolsStatus = new Label { AutoSize = true, ForeColor = Color.Gray, Text = "" };
        layout.Controls.Add(_lblToolsStatus);

        page.Controls.Add(layout);
        return page;
    }

    private void ConvertSelectedCells()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic selection = app.Selection;
            selection.Value2 = selection.Value2;

            _lblToolsStatus.Text = "Selected cells converted to values.";
            _lblToolsStatus.ForeColor = Color.Green;
        }
        catch (Exception ex)
        {
            _lblToolsStatus.Text = $"Error: {ex.Message}";
            _lblToolsStatus.ForeColor = Color.Red;
        }
    }

    private void ConvertAllUseAiCells()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic workbook = app.ActiveWorkbook;

            if (workbook == null)
            {
                _lblToolsStatus.Text = "No active workbook.";
                _lblToolsStatus.ForeColor = Color.Gray;
                return;
            }

            int convertedCount = 0;

            foreach (dynamic sheet in workbook.Worksheets)
            {
                try
                {
                    dynamic usedRange = sheet.UsedRange;
                    if (usedRange == null) continue;

                    foreach (dynamic cell in usedRange)
                    {
                        try
                        {
                            if (cell.HasFormula)
                            {
                                string formula = cell.Formula?.ToString() ?? "";
                                if (formula.IndexOf("USEAI", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    cell.Value2 = cell.Value2;
                                    convertedCount++;
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            _lblToolsStatus.Text = $"Converted {convertedCount} USEAI cell{(convertedCount == 1 ? "" : "s")} to values.";
            _lblToolsStatus.ForeColor = Color.Green;
        }
        catch (Exception ex)
        {
            _lblToolsStatus.Text = $"Error: {ex.Message}";
            _lblToolsStatus.ForeColor = Color.Red;
        }
    }

    // ── About Tab ───────────────────────────────────────────────────

    private TabPage CreateAboutTab()
    {
        var page = new TabPage("About") { Padding = new Padding(16) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5 };

        // App title
        layout.Controls.Add(new Label
        {
            Text = "Copilot Clown — Excel AI Add-in",
            AutoSize = true,
            Font = new Font(Font.FontFamily, 12, FontStyle.Bold),
            Padding = new Padding(0, 0, 0, 8)
        });

        // Version
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        layout.Controls.Add(new Label
        {
            Text = $"Version {version?.ToString(3) ?? "1.0.0"}",
            AutoSize = true,
            ForeColor = Color.Gray,
            Padding = new Padding(0, 0, 0, 12)
        });

        // Developer info
        var infoLabel = new Label
        {
            Text = "Developed by\n\n"
                 + "Ben Idrissa TRAORE, PMP, CISA\n\n"
                 + "Computer System Analyst | Certified Data Analyst\n"
                 + "Specializing in Excel, Power BI, and Business Intelligence\n"
                 + "solutions for health, government, and development\n"
                 + "organizations across West Africa.",
            AutoSize = true,
            Font = new Font(Font.FontFamily, 9),
            Padding = new Padding(0, 0, 0, 12)
        };
        layout.Controls.Add(infoLabel);

        // Contact
        var contactPanel = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, RowCount = 3 };
        contactPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        contactPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        contactPanel.Controls.Add(new Label { Text = "Email:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) }, 0, 0);
        var emailLink = new LinkLabel { Text = "traore.benidrissa@gmail.com", AutoSize = true };
        emailLink.LinkClicked += (s, e) =>
        {
            try { System.Diagnostics.Process.Start("mailto:traore.benidrissa@gmail.com"); } catch { }
        };
        contactPanel.Controls.Add(emailLink, 1, 0);

        contactPanel.Controls.Add(new Label { Text = "Phone:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) }, 0, 1);
        contactPanel.Controls.Add(new Label { Text = "+234 816 026 5792", AutoSize = true }, 1, 1);

        contactPanel.Controls.Add(new Label { Text = "Phone:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) }, 0, 2);
        contactPanel.Controls.Add(new Label { Text = "+226 76 48 64 06", AutoSize = true }, 1, 2);

        layout.Controls.Add(contactPanel);

        // GitHub link
        var ghLink = new LinkLabel { Text = "GitHub: Benidrissa/copilot_clown", AutoSize = true, Padding = new Padding(0, 8, 0, 0) };
        ghLink.LinkClicked += (s, e) =>
        {
            try { System.Diagnostics.Process.Start("https://github.com/Benidrissa/copilot_clown"); } catch { }
        };
        layout.Controls.Add(ghLink);

        page.Controls.Add(layout);
        return page;
    }

    // ── Data helpers ────────────────────────────────────────────────

    private void LoadCurrentSettings()
    {
        var s = _settings.LoadSettings();

        // Keys
        var claudeKey = _settings.GetApiKey(ProviderName.Anthropic);
        if (claudeKey != null) _txtClaudeKey.Text = claudeKey;
        var openaiKey = _settings.GetApiKey(ProviderName.OpenAI);
        if (openaiKey != null) _txtOpenAIKey.Text = openaiKey;

        // Provider & model
        if (s.ActiveProvider == ProviderName.Anthropic)
            _rbClaude.Checked = true;
        else
            _rbOpenAI.Checked = true;
        RefreshModelList();
        _cboModel.SelectedValue = s.ActiveModel;

        // Cache
        _chkCacheEnabled.Checked = s.CacheEnabled;
        _cboCacheTtl.SelectedIndex = s.CacheTtlMinutes switch
        {
            60 => 0,
            360 => 1,
            1440 => 2,
            10080 => 3,
            43200 => 4,
            _ => 2,
        };
        RefreshCacheStats();
    }

    private void RefreshModelList()
    {
        var provider = _rbClaude.Checked ? ProviderName.Anthropic : ProviderName.OpenAI;
        var models = ModelRegistry.GetModels(provider);
        _cboModel.DataSource = models;
        _cboModel.DisplayMember = "DisplayName";
        _cboModel.ValueMember = "Id";
        RefreshModelInfo();
    }

    private void RefreshModelInfo()
    {
        if (_cboModel.SelectedItem is ModelInfo model)
        {
            var ctx = model.ContextWindow >= 1_000_000
                ? $"{model.ContextWindow / 1_000_000.0:F1}M tokens"
                : $"{model.ContextWindow / 1000}K tokens";
            _lblModelInfo.Text = $"Context: {ctx}  |  Pricing: {model.PricingTier}";
        }
    }

    private void RefreshCacheStats()
    {
        var (entries, hits, misses, hitRate) = _cache.GetStats();
        _lblCacheStats.Text = $"Entries: {entries}\nHits: {hits}  |  Misses: {misses}\nHit Rate: {hitRate:P0}";
    }

    private void SaveApiKey(ProviderName provider, string key)
    {
        if (!string.IsNullOrWhiteSpace(key))
        {
            _settings.SetApiKey(provider, key.Trim());
            MessageBox.Show($"{provider} API key saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private async Task TestApiKey(ProviderName provider, string key, Label statusLabel)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            statusLabel.Text = "Enter a key first.";
            statusLabel.ForeColor = Color.Gray;
            return;
        }

        statusLabel.Text = "Testing...";
        statusLabel.ForeColor = Color.Gray;

        try
        {
            var llm = ProviderFactory.Get(provider);
            var result = await llm.CompleteAsync("Say OK", key.Trim(),
                provider == ProviderName.Anthropic ? "claude-3-haiku-20240307" : "gpt-4o-mini",
                maxTokens: 10);
            statusLabel.Text = "Valid";
            statusLabel.ForeColor = Color.Green;
            return;
        }
        catch (Exception ex)
        {
            statusLabel.Text = ex.Message.Length > 120 ? ex.Message.Substring(0, 120) + "..." : ex.Message;
            statusLabel.ForeColor = Color.Red;
        }
    }

    private void SaveModelSelection()
    {
        var s = _settings.LoadSettings();
        s.ActiveProvider = _rbClaude.Checked ? ProviderName.Anthropic : ProviderName.OpenAI;
        if (_cboModel.SelectedItem is ModelInfo model)
            s.ActiveModel = model.Id;
        _settings.SaveSettings(s);
        MessageBox.Show("Model selection saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void SaveCacheSettings()
    {
        var s = _settings.LoadSettings();
        s.CacheEnabled = _chkCacheEnabled.Checked;
        s.CacheTtlMinutes = _cboCacheTtl.SelectedIndex switch
        {
            0 => 60,
            1 => 360,
            2 => 1440,
            3 => 10080,
            4 => 43200,
            _ => 1440,
        };
        _settings.SaveSettings(s);
        MessageBox.Show("Cache settings saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
