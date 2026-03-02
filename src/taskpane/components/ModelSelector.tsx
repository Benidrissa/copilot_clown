import React from "react";
import { ProviderName, getModelsForProvider, ModelInfo } from "../../types";

interface Props {
  activeProvider: ProviderName;
  activeModel: string;
  onProviderChange: (provider: ProviderName) => void;
  onModelChange: (model: string) => void;
}

export function ModelSelector({ activeProvider, activeModel, onProviderChange, onModelChange }: Props) {
  const models = getModelsForProvider(activeProvider);
  const selectedModel = models.find((m) => m.id === activeModel) || models[0];

  const handleProviderSwitch = (provider: ProviderName) => {
    onProviderChange(provider);
    // Auto-select first model of the new provider
    const newModels = getModelsForProvider(provider);
    if (newModels.length > 0) {
      onModelChange(newModels[0].id);
    }
  };

  return (
    <div className="section">
      <div className="section-title">Model Selection</div>

      <div className="provider-toggle">
        <button
          className={activeProvider === "anthropic" ? "active" : ""}
          onClick={() => handleProviderSwitch("anthropic")}
        >
          Claude
        </button>
        <button
          className={activeProvider === "openai" ? "active" : ""}
          onClick={() => handleProviderSwitch("openai")}
        >
          OpenAI
        </button>
      </div>

      <div className="field">
        <label>Model</label>
        <select value={activeModel} onChange={(e) => onModelChange(e.target.value)}>
          {models.map((model) => (
            <option key={model.id} value={model.id}>
              {model.displayName}
            </option>
          ))}
        </select>
        {selectedModel && <ModelInfoDisplay model={selectedModel} />}
      </div>
    </div>
  );
}

function ModelInfoDisplay({ model }: { model: ModelInfo }) {
  const tierLabels: Record<string, string> = {
    low: "Budget",
    medium: "Standard",
    high: "Premium",
  };

  const formatContextWindow = (tokens: number): string => {
    if (tokens >= 1_000_000) return `${(tokens / 1_000_000).toFixed(1)}M tokens`;
    return `${(tokens / 1000).toFixed(0)}K tokens`;
  };

  return (
    <div className="model-info">
      Context: {formatContextWindow(model.contextWindow)} &middot; Pricing: {tierLabels[model.pricingTier]}
    </div>
  );
}
