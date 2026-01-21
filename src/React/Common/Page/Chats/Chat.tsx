import React, { useState } from "react";
import AI_Factory from "../../Function/AI/AI_Factory";
import PII_Filter from "../../Function/PII_Filter";
import ModelSelector, { ModelPath } from "./Model_Selector";

type ChatMode = "normal" | "secure";

interface Message {
  role: "user" | "assistant";
  content: string;
  maskedContent?: string;
}

interface ChatState {
  messages: Message[];
  selectedModel: ModelPath;
}

const Chat: React.FC = () => {
  const [mode, setMode] = useState<ChatMode>("normal");
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [showModelSelector, setShowModelSelector] = useState(false);

  // 일반 채팅과 보안 채팅 상태 분리
  const [normalChat, setNormalChat] = useState<ChatState>({
    messages: [],
    selectedModel: "openai.GPT_4o",
  });
  const [secureChat, setSecureChat] = useState<ChatState>({
    messages: [],
    selectedModel: "openai.GPT_4o",
  });

  const currentChat = mode === "normal" ? normalChat : secureChat;
  const setCurrentChat = mode === "normal" ? setNormalChat : setSecureChat;

  const getModelFunction = (modelPath: ModelPath) => {
    const [provider, model] = modelPath.split(".");
    return (AI_Factory as any)[provider][model].generateText;
  };

  const handleSend = async () => {
    if (!input.trim() || loading) return;

    const userMessage: Message = { role: "user", content: input };
    setCurrentChat((prev) => ({
      ...prev,
      messages: [...prev.messages, userMessage],
    }));
    setInput("");
    setLoading(true);

    try {
      if (mode === "secure") {
        const maskedText = await PII_Filter.getMaskedText(input);
        if (!maskedText || maskedText.trim() === "") {
          throw new Error("PII 필터링 결과가 비어있습니다.");
        }
        const generateText = getModelFunction(secureChat.selectedModel);
        const aiResponse = await generateText(maskedText);
        setSecureChat((prev) => ({
          ...prev,
          messages: [
            ...prev.messages,
            { role: "assistant", content: aiResponse, maskedContent: maskedText },
          ],
        }));
      } else {
        if (!input.trim()) {
          throw new Error("입력 내용이 비어있습니다.");
        }
        const generateText = getModelFunction(normalChat.selectedModel);
        const aiResponse = await generateText(input);
        setNormalChat((prev) => ({
          ...prev,
          messages: [
            ...prev.messages,
            { role: "assistant", content: aiResponse },
          ],
        }));
      }
    } catch (error) {
      setCurrentChat((prev) => ({
        ...prev,
        messages: [
          ...prev.messages,
          { role: "assistant", content: `오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}` },
        ],
      }));
    } finally {
      setLoading(false);
    }
  };

  const handleModelChange = (model: ModelPath) => {
    setCurrentChat((prev) => ({ ...prev, selectedModel: model }));
    setShowModelSelector(false);
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", position: "relative" }}>
      <div style={{ display: "flex", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5" }}>
        <button
          onClick={() => setMode("normal")}
          style={{
            flex: 1,
            padding: "12px",
            border: "none",
            backgroundColor: mode === "normal" ? "#fff" : "transparent",
            cursor: "pointer",
            fontWeight: mode === "normal" ? "600" : "400",
          }}
        >
          일반 채팅
        </button>
        <button
          onClick={() => setMode("secure")}
          style={{
            flex: 1,
            padding: "12px",
            border: "none",
            backgroundColor: mode === "secure" ? "#fff" : "transparent",
            cursor: "pointer",
            fontWeight: mode === "secure" ? "600" : "400",
          }}
        >
          보안 채팅
        </button>
      </div>

      <div style={{ flex: 1, overflowY: "auto", padding: "15px" }}>
        {currentChat.messages.map((msg, idx) => (
          <div key={idx} style={{ marginBottom: "15px" }}>
            <div style={{ fontWeight: "600", marginBottom: "5px", color: msg.role === "user" ? "#2196f3" : "#4caf50" }}>
              {msg.role === "user" ? "사용자" : "AI"}
            </div>
            <div style={{ padding: "10px", backgroundColor: msg.role === "user" ? "#e3f2fd" : "#f1f8e9", borderRadius: "5px" }}>
              {msg.content}
            </div>
            {msg.maskedContent && (
              <div style={{ marginTop: "5px", fontSize: "12px", color: "#666", padding: "5px", backgroundColor: "#fff3cd", borderRadius: "3px" }}>
                마스킹된 요청: {msg.maskedContent}
              </div>
            )}
          </div>
        ))}
        {loading && <div style={{ color: "#666" }}>응답 중...</div>}
      </div>

      <div style={{ padding: "15px", borderTop: "1px solid #ddd", position: "relative" }}>
        {showModelSelector && (
          <ModelSelector
            selectedModel={currentChat.selectedModel}
            onModelChange={handleModelChange}
          />
        )}
        <div style={{ marginBottom: "8px", fontSize: "12px", color: "#666" }}>
          현재 모델: {currentChat.selectedModel.split(".")[1]}
          <button
            onClick={() => setShowModelSelector(!showModelSelector)}
            style={{
              marginLeft: "10px",
              padding: "4px 8px",
              fontSize: "11px",
              border: "1px solid #ddd",
              borderRadius: "3px",
              backgroundColor: "#fff",
              cursor: "pointer",
            }}
          >
            {showModelSelector ? "닫기" : "변경"}
          </button>
        </div>
        <div style={{ display: "flex", gap: "10px" }}>
          <input
            type="text"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyPress={(e) => e.key === "Enter" && handleSend()}
            placeholder="메시지를 입력하세요..."
            style={{ flex: 1, padding: "10px", border: "1px solid #ddd", borderRadius: "5px" }}
          />
          <button
            onClick={handleSend}
            disabled={loading}
            style={{
              padding: "10px 20px",
              backgroundColor: "#2196f3",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: loading ? "not-allowed" : "pointer",
              opacity: loading ? 0.6 : 1,
            }}
          >
            전송
          </button>
        </div>
      </div>
    </div>
  );
};

export default Chat;
