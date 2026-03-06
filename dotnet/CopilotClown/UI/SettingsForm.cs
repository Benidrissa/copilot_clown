using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
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
