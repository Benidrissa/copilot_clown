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

    public SettingsService()
    {
        Directory.CreateDirectory(AppDataDir);
    }

    // ── Settings ────────────────────────────────────────────────────

    public AppSettings LoadSettings()
    {
        if (!File.Exists(SettingsFile))
            return new AppSettings();

        try
        {
            var json = File.ReadAllText(SettingsFile);
            return JsonHelper.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void SaveSettings(AppSettings settings)
    {
        var json = JsonHelper.Serialize(settings);
        File.WriteAllText(SettingsFile, json);
    }

    // ── API Keys (DPAPI encrypted) ─────────────────────────────────

    public string GetApiKey(ProviderName provider)
    {
        var keys = LoadKeys();
        var key = provider.ToString();
        return keys.TryGetValue(key, out var encrypted) ? Unprotect(encrypted) : null;
    }

    public void SetApiKey(ProviderName provider, string apiKey)
    {
        var keys = LoadKeys();
        keys[provider.ToString()] = Protect(apiKey);
        SaveKeys(keys);
    }

    public void RemoveApiKey(ProviderName provider)
    {
        var keys = LoadKeys();
        keys.Remove(provider.ToString());
        SaveKeys(keys);
    }

    public bool HasApiKey(ProviderName provider)
    {
        var keys = LoadKeys();
        return keys.ContainsKey(provider.ToString());
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
