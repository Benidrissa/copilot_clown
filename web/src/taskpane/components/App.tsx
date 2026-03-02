import React, { useState, useCallback } from "react";
import { ApiKeySettings } from "./ApiKeySettings";
import { ModelSelector } from "./ModelSelector";
import { CacheManager } from "./CacheManager";
import { ProviderName } from "../../types";
import "../taskpane.css";

export function App() {
  const services = window.aillmServices;
  const settings = services.settingsService.getSettings();

  const [activeProvider, setActiveProvider] = useState<ProviderName>(settings.activeProvider);
  const [activeModel, setActiveModel] = useState(settings.activeModel);

  const handleProviderChange = useCallback(
    (provider: ProviderName) => {
      setActiveProvider(provider);
      services.settingsService.saveSettings({ activeProvider: provider });
    },
    [services.settingsService],
  );

  const handleModelChange = useCallback(
    (model: string) => {
      setActiveModel(model);
      services.settingsService.saveSettings({ activeModel: model });
    },
    [services.settingsService],
  );

  return (
    <div className="app">
      <div className="app-header">
        <h1>Copilot Clown</h1>
      </div>

      <ApiKeySettings settingsService={services.settingsService} />

      <div className="divider" />

      <ModelSelector
        activeProvider={activeProvider}
        activeModel={activeModel}
        onProviderChange={handleProviderChange}
        onModelChange={handleModelChange}
      />

      <div className="divider" />

      <CacheManager
        cacheService={services.cacheService}
        settingsService={services.settingsService}
      />
    </div>
  );
}
