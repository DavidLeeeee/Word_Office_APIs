import React from "react";

export type ModelPath = 
  | "openai.GPT_4o"
  | "openai.GPT_4_Turbo"
  | "openai.GPT_4"
  | "openai.GPT_3_5_Turbo";

interface ModelSelectorProps {
  selectedModel: ModelPath;
  onModelChange: (model: ModelPath) => void;
}

const ModelSelector: React.FC<ModelSelectorProps> = ({ selectedModel, onModelChange }) => {
  const models: { value: ModelPath; label: string }[] = [
    { value: "openai.GPT_4o", label: "GPT-4o" },
    { value: "openai.GPT_4_Turbo", label: "GPT-4 Turbo" },
    { value: "openai.GPT_4", label: "GPT-4" },
    { value: "openai.GPT_3_5_Turbo", label: "GPT-3.5 Turbo" },
  ];

  return (
    <div style={{
      position: "absolute",
      bottom: "70px",
      left: "15px",
      right: "15px",
      backgroundColor: "#fff",
      border: "1px solid #ddd",
      borderRadius: "8px",
      padding: "10px",
      boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
      zIndex: 100,
    }}>
      <div style={{ fontSize: "12px", color: "#666", marginBottom: "8px" }}>모델 선택</div>
      <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
        {models.map((model) => (
          <button
            key={model.value}
            onClick={() => onModelChange(model.value)}
            style={{
              padding: "6px 12px",
              border: selectedModel === model.value ? "2px solid #2196f3" : "1px solid #ddd",
              borderRadius: "5px",
              backgroundColor: selectedModel === model.value ? "#e3f2fd" : "#fff",
              cursor: "pointer",
              fontSize: "12px",
              fontWeight: selectedModel === model.value ? "600" : "400",
            }}
          >
            {model.label}
          </button>
        ))}
      </div>
    </div>
  );
};

export default ModelSelector;
