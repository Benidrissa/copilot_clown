using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class SettingsService
{
    private static readonly string AppDataDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CopilotClown");
    private static readonly string SettingsFile = Path.Combine(AppDataDir, "settings.json");
    private static readonly string KeyFile = Path.Combine(AppDataDir, "keys.dat");
    private static readonly string PromptsFile = Path.Combine(AppDataDir, "prompts.json");

    // In-memory caches — avoid disk I/O on every cell calculation
    private AppSettings _cachedSettings;
    private Dictionary<ProviderName, string> _cachedKeys = new Dictionary<ProviderName, string>();
    private DateTime _settingsCacheTime;
    private DateTime _keysCacheTime;
    private const int CacheSeconds = 5; // Re-read from disk at most every 5 seconds

    public SettingsService()
    {
        Directory.CreateDirectory(AppDataDir);
    }

    // ── Settings (cached) ───────────────────────────────────────────

    public AppSettings LoadSettings()
    {
        if (_cachedSettings != null && (DateTime.UtcNow - _settingsCacheTime).TotalSeconds < CacheSeconds)
            return _cachedSettings;

        if (!File.Exists(SettingsFile))
        {
            _cachedSettings = new AppSettings();
            _settingsCacheTime = DateTime.UtcNow;
            return _cachedSettings;
        }

        try
        {
            var json = File.ReadAllText(SettingsFile);
            _cachedSettings = JsonHelper.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            _cachedSettings = new AppSettings();
        }
        _settingsCacheTime = DateTime.UtcNow;
        return _cachedSettings;
    }

    public void SaveSettings(AppSettings settings)
    {
        var json = JsonHelper.Serialize(settings);
        File.WriteAllText(SettingsFile, json);
        _cachedSettings = settings;
        _settingsCacheTime = DateTime.UtcNow;
    }

    // ── API Keys (cached + DPAPI encrypted on disk) ─────────────────

    public string GetApiKey(ProviderName provider)
    {
        if (_cachedKeys.TryGetValue(provider, out var cached) &&
            (DateTime.UtcNow - _keysCacheTime).TotalSeconds < CacheSeconds)
            return cached;

        var keys = LoadKeys();
        var key = provider.ToString();
        var result = keys.TryGetValue(key, out var encrypted) ? Unprotect(encrypted) : null;
        _cachedKeys[provider] = result;
        _keysCacheTime = DateTime.UtcNow;
        return result;
    }

    public void SetApiKey(ProviderName provider, string apiKey)
    {
        var keys = LoadKeys();
        keys[provider.ToString()] = Protect(apiKey);
        SaveKeys(keys);
        _cachedKeys[provider] = apiKey;
        _keysCacheTime = DateTime.UtcNow;
    }

    public void RemoveApiKey(ProviderName provider)
    {
        var keys = LoadKeys();
        keys.Remove(provider.ToString());
        SaveKeys(keys);
        _cachedKeys.Remove(provider);
    }

    public bool HasApiKey(ProviderName provider)
    {
        return GetApiKey(provider) != null;
    }

    // ── Prompt Library ────────────────────────────────────────────

    public List<SystemPromptEntry> LoadPrompts()
    {
        if (!File.Exists(PromptsFile))
            return new List<SystemPromptEntry>();

        try
        {
            var json = File.ReadAllText(PromptsFile);
            return JsonHelper.Deserialize<List<SystemPromptEntry>>(json)
                   ?? new List<SystemPromptEntry>();
        }
        catch
        {
            return new List<SystemPromptEntry>();
        }
    }

    public void SavePrompts(List<SystemPromptEntry> prompts)
    {
        var json = JsonHelper.Serialize(prompts);
        File.WriteAllText(PromptsFile, json);
    }

    // ── DPAPI helpers ───────────────────────────────────────────────

    private static string Protect(string plainText)
    {
        var bytes = Encoding.UTF8.GetBytes(plainText);
        var encrypted = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(encrypted);
    }

    private static string Unprotect(string encryptedBase64)
    {
        try
        {
            var encrypted = Convert.FromBase64String(encryptedBase64);
            var bytes = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(bytes);
        }
        catch
        {
            return null;
        }
    }

    private Dictionary<string, string> LoadKeys()
    {
        if (!File.Exists(KeyFile))
            return new Dictionary<string, string>();

        try
        {
            var json = File.ReadAllText(KeyFile);
            return JsonHelper.Deserialize<Dictionary<string, string>>(json)
                   ?? new Dictionary<string, string>();
        }
        catch
        {
            return new Dictionary<string, string>();
        }
    }

    private void SaveKeys(Dictionary<string, string> keys)
    {
        var json = JsonHelper.Serialize(keys);
        File.WriteAllText(KeyFile, json);
    }
}
