import React, { useState, useCallback } from "react";
import { SettingsService } from "../../services/SettingsService";
import { getProvider } from "../../services/providers/LLMProvider";
import { ProviderName } from "../../types";

interface Props {
  settingsService: SettingsService;
}

type KeyStatus = "unknown" | "valid" | "invalid" | "testing";

export function ApiKeySettings({ settingsService }: Props) {
  return (
    <div className="section">
      <div className="section-title">API Keys</div>
      <ProviderKeyInput provider="anthropic" label="Anthropic (Claude)" settingsService={settingsService} />
      <ProviderKeyInput provider="openai" label="OpenAI" settingsService={settingsService} />
    </div>
  );
}

function ProviderKeyInput({
  provider,
  label,
  settingsService,
}: {
  provider: ProviderName;
  label: string;
  settingsService: SettingsService;
}) {
  const existingKey = settingsService.getApiKey(provider);
  const [key, setKey] = useState(existingKey || "");
  const [status, setStatus] = useState<KeyStatus>(existingKey ? "unknown" : "unknown");
  const [saved, setSaved] = useState(!!existingKey);

  const handleSave = useCallback(() => {
    if (key.trim()) {
      settingsService.setApiKey(provider, key.trim());
      setSaved(true);
      setStatus("unknown");
    }
  }, [key, provider, settingsService]);

  const handleTest = useCallback(async () => {
    if (!key.trim()) return;

    // Save first if not saved
    if (!saved) {
      settingsService.setApiKey(provider, key.trim());
      setSaved(true);
    }

    setStatus("testing");
    try {
      const providerImpl = getProvider(provider);
      const valid = await providerImpl.validateKey(key.trim());
      setStatus(valid ? "valid" : "invalid");
    } catch {
      setStatus("invalid");
    }
  }, [key, provider, saved, settingsService]);

  const handleClear = useCallback(() => {
    settingsService.removeApiKey(provider);
    setKey("");
    setSaved(false);
    setStatus("unknown");
  }, [provider, settingsService]);

  return (
    <div className="field">
      <label>{label}</label>
      <input
        type="password"
        value={key}
        onChange={(e) => {
          setKey(e.target.value);
          setSaved(false);
          setStatus("unknown");
        }}
        placeholder={`Enter ${label} API key`}
      />
      <div className="btn-row">
        <button className="btn btn-primary btn-sm" onClick={handleSave} disabled={!key.trim()}>
          Save
        </button>
        <button className="btn btn-secondary btn-sm" onClick={handleTest} disabled={!key.trim()}>
          {status === "testing" ? "Testing..." : "Test"}
        </button>
        {saved && (
          <button className="btn btn-secondary btn-sm" onClick={handleClear}>
            Clear
          </button>
        )}
      </div>
      {status !== "unknown" && status !== "testing" && (
        <span className={`status-badge status-${status}`} style={{ marginTop: 4 }}>
          {status === "valid" ? "Valid" : "Invalid"}
        </span>
      )}
    </div>
  );
}
