using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
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
    private TextBox _txtGeminiKey;
    private Label _lblClaudeStatus;
    private Label _lblOpenAIStatus;
    private Label _lblGeminiStatus;

    // Model tab
    private RadioButton _rbClaude;
    private RadioButton _rbOpenAI;
    private RadioButton _rbGemini;
    private ComboBox _cboModel;
    private Label _lblModelInfo;

    // Cache tab
    private Label _lblCacheStats;
    private ComboBox _cboCacheTtl;
    private CheckBox _chkCacheEnabled;

    // Tools tab
    private Label _lblToolsStatus;

    // System Prompt tab
    private ComboBox _cboPromptLibrary;
    private TextBox _txtPromptName;
    private TextBox _txtSystemPrompt;
    private Label _lblSystemPromptCount;
    private List<SystemPromptEntry> _prompts;

    // Parameters tab
    private TrackBar _trkTemperature;
    private Label _lblTemperatureValue;
    private NumericUpDown _nudMaxTokens;
    private TrackBar _trkTopP;
    private Label _lblTopPValue;

    // Rate Limits tab
    private Label _lblRateLimitInfo;
    private Timer _rateLimitTimer;

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
        Size = new Size(520, 540);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;

        var tabs = new TabControl { Dock = DockStyle.Fill };

        tabs.TabPages.Add(CreateApiKeysTab());
        tabs.TabPages.Add(CreateModelTab());
        tabs.TabPages.Add(CreateSystemPromptTab());
        tabs.TabPages.Add(CreateParametersTab());
        tabs.TabPages.Add(CreateCacheTab());
        tabs.TabPages.Add(CreateRateLimitsTab());
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

        // Spacer
        layout.Controls.Add(new Label { Height = 16 });

        // Google Gemini section
        layout.Controls.Add(new Label { Text = "Google (Gemini) API Key:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _txtGeminiKey = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
        layout.Controls.Add(_txtGeminiKey);

        var geminiBtnPanel = new FlowLayoutPanel { AutoSize = true };
        var btnSaveGemini = new Button { Text = "Save", Width = 70 };
        btnSaveGemini.Click += (s, e) => SaveApiKey(ProviderName.Google, _txtGeminiKey.Text);
        var btnTestGemini = new Button { Text = "Test", Width = 70 };
        btnTestGemini.Click += async (s, e) => await TestApiKey(ProviderName.Google, _txtGeminiKey.Text, _lblGeminiStatus);
        geminiBtnPanel.Controls.AddRange(new Control[] { btnSaveGemini, btnTestGemini });
        layout.Controls.Add(geminiBtnPanel);
        _lblGeminiStatus = new Label { AutoSize = true, ForeColor = Color.Gray, Text = "" };
        layout.Controls.Add(_lblGeminiStatus);

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
        _rbGemini = new RadioButton { Text = "Google Gemini", AutoSize = true };
        _rbClaude.CheckedChanged += (s, e) => RefreshModelList();
        _rbOpenAI.CheckedChanged += (s, e) => RefreshModelList();
        _rbGemini.CheckedChanged += (s, e) => RefreshModelList();
        providerPanel.Controls.AddRange(new Control[] { _rbClaude, _rbOpenAI, _rbGemini });
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
        var btnClearWb = new Button { Text = "Clear Workbook Cache", Width = 140, ForeColor = Color.DarkRed };
        btnClearWb.Click += (s, e) =>
        {
            WorkbookCache.Clear();
            RefreshCacheStats();
            MessageBox.Show("Workbook embedded cache cleared.", "Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
        var btnClearDisk = new Button { Text = "Clear Disk Cache", Width = 120, ForeColor = Color.DarkRed };
        btnClearDisk.Click += (s, e) =>
        {
            UseAiFunction.DiskCacheInstance.Clear();
            RefreshCacheStats();
            MessageBox.Show("Disk cache cleared.", "Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
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
        btnPanel.Controls.AddRange(new Control[] { btnClear, btnClearWb, btnClearDisk, btnClearFiles, btnSaveCache });
        layout.Controls.Add(btnPanel);

        page.Controls.Add(layout);
        return page;
    }

    // ── System Prompt Tab ──────────────────────────────────────────

    private TabPage CreateSystemPromptTab()
    {
        var page = new TabPage("System Prompt") { Padding = new Padding(12) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 8 };

        // Saved prompts dropdown
        layout.Controls.Add(new Label
        {
            Text = "Saved Prompts:",
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        });

        var selectorPanel = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        _cboPromptLibrary = new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDownList };
        _cboPromptLibrary.SelectedIndexChanged += (s, e) => OnPromptSelected();
        selectorPanel.Controls.Add(_cboPromptLibrary);

        var btnLoad = new Button { Text = "Load", Width = 55 };
        btnLoad.Click += (s, e) => OnPromptSelected();
        selectorPanel.Controls.Add(btnLoad);

        var btnDelete = new Button { Text = "Delete", Width = 55, ForeColor = Color.DarkRed };
        btnDelete.Click += (s, e) => DeleteSelectedPrompt();
        selectorPanel.Controls.Add(btnDelete);
        layout.Controls.Add(selectorPanel);

        // Prompt name
        layout.Controls.Add(new Label { Text = "Prompt Name:", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _txtPromptName = new TextBox { Dock = DockStyle.Fill, MaxLength = 100 };
        layout.Controls.Add(_txtPromptName);

        // Prompt text
        layout.Controls.Add(new Label
        {
            Text = "System prompt sent with every API call (persona / global instructions):",
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        });

        _txtSystemPrompt = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Height = 160,
            MaxLength = 10000
        };
        _txtSystemPrompt.TextChanged += (s, e) =>
        {
            _lblSystemPromptCount.Text = $"{_txtSystemPrompt.Text.Length} / 10,000 characters";
        };
        layout.Controls.Add(_txtSystemPrompt);

        _lblSystemPromptCount = new Label { AutoSize = true, ForeColor = Color.Gray, Text = "0 / 10,000 characters" };
        layout.Controls.Add(_lblSystemPromptCount);

        var btnPanel = new FlowLayoutPanel { AutoSize = true };
        var btnSaveAs = new Button { Text = "Save to Library", Width = 110 };
        btnSaveAs.Click += (s, e) => SavePromptToLibrary();
        var btnActivate = new Button { Text = "Set as Active", Width = 100 };
        btnActivate.Click += (s, e) => SaveSystemPrompt();
        var btnClear = new Button { Text = "Clear", Width = 55 };
        btnClear.Click += (s, e) => { _txtSystemPrompt.Text = ""; _txtPromptName.Text = ""; };
        btnPanel.Controls.AddRange(new Control[] { btnSaveAs, btnActivate, btnClear });
        layout.Controls.Add(btnPanel);

        page.Controls.Add(layout);
        return page;
    }

    private void OnPromptSelected()
    {
        if (_cboPromptLibrary.SelectedItem is SystemPromptEntry entry)
        {
            _txtPromptName.Text = entry.Name;
            _txtSystemPrompt.Text = entry.Text;
        }
    }

    private void SavePromptToLibrary()
    {
        var name = _txtPromptName.Text.Trim();
        if (string.IsNullOrEmpty(name))
        {
            MessageBox.Show("Enter a name for this prompt.", "Name Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var existing = _prompts.FindIndex(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (existing >= 0)
        {
            _prompts[existing].Text = _txtSystemPrompt.Text;
        }
        else
        {
            _prompts.Add(new SystemPromptEntry(name, _txtSystemPrompt.Text));
        }

        _settings.SavePrompts(_prompts);
        RefreshPromptList(name);
        MessageBox.Show($"Prompt \"{name}\" saved to library.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void DeleteSelectedPrompt()
    {
        if (_cboPromptLibrary.SelectedItem is not SystemPromptEntry entry) return;

        var result = MessageBox.Show(
            $"Delete \"{entry.Name}\" from the prompt library?",
            "Delete Prompt",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);
        if (result != DialogResult.Yes) return;

        _prompts.RemoveAll(p => p.Name == entry.Name);
        _settings.SavePrompts(_prompts);
        _txtPromptName.Text = "";
        _txtSystemPrompt.Text = "";
        RefreshPromptList(null);
    }

    private void RefreshPromptList(string selectName)
    {
        _cboPromptLibrary.Items.Clear();
        _cboPromptLibrary.Items.Add("(none)");
        foreach (var p in _prompts)
            _cboPromptLibrary.Items.Add(p);

        if (selectName != null)
        {
            for (int i = 1; i < _cboPromptLibrary.Items.Count; i++)
            {
                if (_cboPromptLibrary.Items[i] is SystemPromptEntry e && e.Name == selectName)
                {
                    _cboPromptLibrary.SelectedIndex = i;
                    return;
                }
            }
        }
        _cboPromptLibrary.SelectedIndex = 0;
    }

    // ── Parameters Tab ─────────────────────────────────────────────

    private TabPage CreateParametersTab()
    {
        var page = new TabPage("Parameters") { Padding = new Padding(12) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 10 };

        // Temperature
        layout.Controls.Add(new Label { Text = "Temperature (0.0 = deterministic, 2.0 = creative):", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _trkTemperature = new TrackBar { Minimum = 0, Maximum = 200, TickFrequency = 10, Dock = DockStyle.Fill };
        _lblTemperatureValue = new Label { AutoSize = true, ForeColor = Color.Gray };
        _trkTemperature.ValueChanged += (s, e) =>
        {
            _lblTemperatureValue.Text = $"{_trkTemperature.Value / 100.0:F2}";
        };
        layout.Controls.Add(_trkTemperature);
        layout.Controls.Add(_lblTemperatureValue);

        // Max Tokens
        layout.Controls.Add(new Label { Text = "Max Tokens (response length):", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _nudMaxTokens = new NumericUpDown { Minimum = 1, Maximum = 32768, Dock = DockStyle.Fill };
        layout.Controls.Add(_nudMaxTokens);

        // Top P
        layout.Controls.Add(new Label { Text = "Top P (nucleus sampling, 0.0 – 1.0):", AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
        _trkTopP = new TrackBar { Minimum = 0, Maximum = 100, TickFrequency = 5, Dock = DockStyle.Fill };
        _lblTopPValue = new Label { AutoSize = true, ForeColor = Color.Gray };
        _trkTopP.ValueChanged += (s, e) =>
        {
            _lblTopPValue.Text = $"{_trkTopP.Value / 100.0:F2}";
        };
        layout.Controls.Add(_trkTopP);
        layout.Controls.Add(_lblTopPValue);

        var btnSave = new Button { Text = "Save Parameters", Width = 130 };
        btnSave.Click += (s, e) => SaveParameters();
        layout.Controls.Add(btnSave);

        page.Controls.Add(layout);
        return page;
    }

    // ── Rate Limits Tab ────────────────────────────────────────────

    private TabPage CreateRateLimitsTab()
    {
        var page = new TabPage("Rate Limits") { Padding = new Padding(12) };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };

        layout.Controls.Add(new Label
        {
            Text = "Per-Provider Rate Limit Status (refreshes every 5s):",
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        });

        _lblRateLimitInfo = new Label
        {
            AutoSize = true,
            Font = new Font("Consolas", 9),
            Text = "Loading..."
        };
        layout.Controls.Add(_lblRateLimitInfo);

        var btnReset = new Button { Text = "Reset All Counters", Width = 140 };
        btnReset.Click += (s, e) =>
        {
            foreach (var kv in UseAiFunction.AllRateLimiters)
                kv.Value.Reset();
            RefreshRateLimitInfo();
        };
        layout.Controls.Add(btnReset);

        // Auto-refresh timer
        _rateLimitTimer = new Timer { Interval = 5000 };
        _rateLimitTimer.Tick += (s, e) => RefreshRateLimitInfo();
        _rateLimitTimer.Start();

        page.Controls.Add(layout);
        RefreshRateLimitInfo();
        return page;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _rateLimitTimer?.Stop();
        _rateLimitTimer?.Dispose();
        base.OnFormClosed(e);
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

            int count = ConvertUndoState.SavedFormulas.Count;
            _lblToolsStatus.Text = $"Converted {count} cell(s) to values. Use 'Undo Convert' in the ribbon to restore.";
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

            // Save formulas for undo
            ConvertUndoState.Clear();
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
                                    string sheetName = cell.Worksheet.Name;
                                    string address = cell.Address;
                                    ConvertUndoState.SavedFormulas[$"{sheetName}!{address}"] = formula;

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

            _lblToolsStatus.Text = $"Converted {convertedCount} USEAI cell{(convertedCount == 1 ? "" : "s")} to values. Use 'Undo Convert' in the ribbon to restore.";
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
        var geminiKey = _settings.GetApiKey(ProviderName.Google);
        if (geminiKey != null) _txtGeminiKey.Text = geminiKey;

        // Provider & model
        if (s.ActiveProvider == ProviderName.Anthropic)
            _rbClaude.Checked = true;
        else if (s.ActiveProvider == ProviderName.Google)
            _rbGemini.Checked = true;
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

        // System prompt
        _prompts = _settings.LoadPrompts();
        _txtSystemPrompt.Text = s.SystemPrompt ?? "";
        _txtPromptName.Text = s.ActivePromptName ?? "";
        _lblSystemPromptCount.Text = $"{_txtSystemPrompt.Text.Length} / 10,000 characters";
        RefreshPromptList(s.ActivePromptName);

        // Parameters
        _trkTemperature.Value = Math.Max(0, Math.Min(200, (int)(s.Temperature * 100)));
        _lblTemperatureValue.Text = $"{s.Temperature:F2}";
        _nudMaxTokens.Value = Math.Max(1, Math.Min(32768, s.MaxTokens));
        _trkTopP.Value = Math.Max(0, Math.Min(100, (int)(s.TopP * 100)));
        _lblTopPValue.Text = $"{s.TopP:F2}";
    }

    private void RefreshModelList()
    {
        ProviderName provider;
        if (_rbClaude.Checked)
            provider = ProviderName.Anthropic;
        else if (_rbGemini.Checked)
            provider = ProviderName.Google;
        else
            provider = ProviderName.OpenAI;
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
        var wbCount = WorkbookCache.Count;
        var diskCount = UseAiFunction.DiskCacheInstance.Count;
        _lblCacheStats.Text =
            $"Memory Cache:   {entries} entries  |  Hits: {hits}  Misses: {misses}  ({hitRate:P0})\n" +
            $"Workbook Cache: {wbCount} entries  (embedded in .xlsx, travels with file)\n" +
            $"Disk Cache:     {diskCount} entries  (survives Excel restarts)";
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
            string testModel;
            switch (provider)
            {
                case ProviderName.Anthropic: testModel = "claude-3-haiku-20240307"; break;
                case ProviderName.Google: testModel = "gemini-2.0-flash-lite"; break;
                default: testModel = "gpt-4o-mini"; break;
            }
            var result = await llm.CompleteAsync("Say OK", key.Trim(), testModel, maxTokens: 10);
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
        if (_rbClaude.Checked)
            s.ActiveProvider = ProviderName.Anthropic;
        else if (_rbGemini.Checked)
            s.ActiveProvider = ProviderName.Google;
        else
            s.ActiveProvider = ProviderName.OpenAI;
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

    private void SaveSystemPrompt()
    {
        var s = _settings.LoadSettings();
        s.SystemPrompt = _txtSystemPrompt.Text;
        s.ActivePromptName = _txtPromptName.Text.Trim();
        _settings.SaveSettings(s);
        MessageBox.Show("System prompt set as active.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void SaveParameters()
    {
        var s = _settings.LoadSettings();
        s.Temperature = _trkTemperature.Value / 100.0;
        s.MaxTokens = (int)_nudMaxTokens.Value;
        s.TopP = _trkTopP.Value / 100.0;
        _settings.SaveSettings(s);
        MessageBox.Show("Parameters saved.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void RefreshRateLimitInfo()
    {
        var sb = new System.Text.StringBuilder();
        var providers = new[] { ProviderName.Anthropic, ProviderName.OpenAI, ProviderName.Google };

        foreach (var provider in providers)
        {
            var limiter = UseAiFunction.GetRateLimiter(provider);
            var remaining = limiter.RemainingCalls;
            var max = limiter.MaxCalls;
            var usage = limiter.UsagePercentage;

            string status;
            if (limiter.IsLimited)
            {
                var wait = limiter.FormatWaitTime() ?? "<1s";
                status = $"\u274c Rate Limited (wait {wait})";
            }
            else if (limiter.IsNearLimit)
            {
                status = "\u26a0\ufe0f Near Limit";
            }
            else
            {
                status = "\u2705 Available";
            }

            sb.AppendLine($"{provider}:");
            sb.AppendLine($"  {remaining} / {max} calls remaining  [{usage:F0}% used]");
            sb.AppendLine($"  Status: {status}");
            sb.AppendLine();
        }

        // Model availability matrix
        var settings = _settings.LoadSettings();
        sb.AppendLine("───── Model Availability ─────");
        foreach (var provider in providers)
        {
            var limiter = UseAiFunction.GetRateLimiter(provider);
            var models = ModelRegistry.GetModels(provider);
            string icon = limiter.IsLimited ? "\u274c" : limiter.IsNearLimit ? "\u26a0\ufe0f" : "\u2705";

            foreach (var m in models)
            {
                var active = m.Id == settings.ActiveModel ? " *" : "";
                var wait = limiter.IsLimited ? $" (wait {limiter.FormatWaitTime() ?? "<1s"})" : "";
                sb.AppendLine($"  {icon} {m.DisplayName}{active}{wait}");
            }
        }

        _lblRateLimitInfo.Text = sb.ToString();
    }
}
