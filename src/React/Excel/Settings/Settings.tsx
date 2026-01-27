import React, { useState } from "react";

/* global Excel */

const Settings: React.FC = () => {
  const [result, setResult] = useState("");
  const [settingKey, setSettingKey] = useState("");
  const [settingValue, setSettingValue] = useState("");
  const [valueType, setValueType] = useState<"string" | "number" | "boolean">("string");

  // 1. 설정 목록 가져오기
  const listSettings = async () => {
    try {
      await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        settings.load("items/key,items/value");
        await context.sync();

        if (settings.items.length === 0) {
          setResult("현재 워크북에 설정이 없습니다.");
          return;
        }

        let resultText = `설정 목록 (${settings.items.length}개):\n\n`;
        settings.items.forEach((setting, index) => {
          resultText += `${index + 1}. ${setting.key}\n`;
          resultText += `   값: ${JSON.stringify(setting.value)}\n`;
          resultText += `   타입: ${typeof setting.value}\n\n`;
        });

        resultText += `과정:\n1. context.workbook.settings로 설정 컬렉션 가져오기\n2. settings.load("items/key,items/value")로 속성 로드\n3. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 설정 추가/업데이트
  const addSetting = async () => {
    if (!settingKey.trim()) {
      setResult("설정 키를 입력해주세요.");
      return;
    }

    if (!settingValue.trim() && valueType !== "boolean") {
      setResult("설정 값을 입력해주세요.");
      return;
    }

    try {
      let value: any;
      if (valueType === "number") {
        value = parseFloat(settingValue);
        if (isNaN(value)) {
          setResult("숫자 형식이 올바르지 않습니다.");
          return;
        }
      } else if (valueType === "boolean") {
        value = settingValue.toLowerCase() === "true" || settingValue === "1";
      } else {
        value = settingValue;
      }

      await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const setting = settings.add(settingKey, value);
        setting.load("key,value");
        await context.sync();

        setResult(`설정 추가/업데이트 완료!\n키: ${setting.key}\n값: ${JSON.stringify(setting.value)}\n타입: ${typeof setting.value}\n\n과정:\n1. context.workbook.settings.add("${settingKey}", ${JSON.stringify(value)})로 설정 추가\n2. setting.load()로 속성 로드\n3. context.sync()로 동기화`);
        setSettingKey("");
        setSettingValue("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 3. 설정 값 읽기
  const getSetting = async () => {
    if (!settingKey.trim()) {
      setResult("설정 키를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const setting = settings.getItem(settingKey);
        setting.load("key,value");
        await context.sync();

        setResult(`설정 읽기 완료!\n키: ${setting.key}\n값: ${JSON.stringify(setting.value)}\n타입: ${typeof setting.value}\n\n과정:\n1. context.workbook.settings.getItem("${settingKey}")로 설정 가져오기\n2. setting.load("key,value")로 속성 로드\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 설정 키를 확인해주세요.`);
    }
  };

  // 4. 설정 값 업데이트
  const updateSetting = async () => {
    if (!settingKey.trim()) {
      setResult("설정 키를 입력해주세요.");
      return;
    }

    if (!settingValue.trim() && valueType !== "boolean") {
      setResult("설정 값을 입력해주세요.");
      return;
    }

    try {
      let value: any;
      if (valueType === "number") {
        value = parseFloat(settingValue);
        if (isNaN(value)) {
          setResult("숫자 형식이 올바르지 않습니다.");
          return;
        }
      } else if (valueType === "boolean") {
        value = settingValue.toLowerCase() === "true" || settingValue === "1";
      } else {
        value = settingValue;
      }

      await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const setting = settings.getItem(settingKey);
        setting.load("key,value");
        await context.sync();

        const oldValue = setting.value;
        setting.value = value;
        await context.sync();

        setResult(`설정 값 업데이트 완료!\n키: ${setting.key}\n이전 값: ${JSON.stringify(oldValue)}\n새 값: ${JSON.stringify(setting.value)}\n\n과정:\n1. context.workbook.settings.getItem("${settingKey}")로 설정 가져오기\n2. setting.value = ${JSON.stringify(value)}로 값 업데이트\n3. context.sync()로 동기화`);
        setSettingKey("");
        setSettingValue("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 5. 설정 삭제
  const deleteSetting = async () => {
    if (!settingKey.trim()) {
      setResult("삭제할 설정 키를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        const setting = settings.getItem(settingKey);
        setting.load("key");
        await context.sync();

        const deletedKey = setting.key;
        setting.delete();
        await context.sync();

        setResult(`설정 삭제 완료!\n삭제된 설정: ${deletedKey}\n\n과정:\n1. context.workbook.settings.getItem("${settingKey}")로 설정 가져오기\n2. setting.delete()로 설정 삭제\n3. context.sync()로 동기화`);
        setSettingKey("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 설정</h3>

        {/* 안내 섹션 */}
        <div style={{
          marginBottom: "20px",
          padding: "15px",
          backgroundColor: "#e3f2fd",
          borderRadius: "5px",
          border: "1px solid #2196f3",
          fontSize: "13px",
          lineHeight: "1.6"
        }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>⚙️ Excel 설정 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 설정은 워크북에 Add-in 전용 키-값 쌍을 저장하는 기능입니다. 이 설정은 워크북과 함께 저장되며, Add-in이 다시 로드될 때 유지됩니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>설정 추가/업데이트 (문자열, 숫자, 불린)</li>
            <li>설정 목록 조회</li>
            <li>설정 값 읽기</li>
            <li>설정 값 업데이트</li>
            <li>설정 삭제</li>
          </ul>
        </div>

        {/* 설정 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 설정 목록</h4>
          <button
            onClick={listSettings}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            설정 목록 보기
          </button>
        </div>

        {/* 설정 추가/업데이트 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 설정 추가/업데이트</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>설정 키:</label>
            <input
              type="text"
              value={settingKey}
              onChange={(e) => setSettingKey(e.target.value)}
              placeholder="예: userPreference"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>값 타입:</label>
            <select
              value={valueType}
              onChange={(e) => setValueType(e.target.value as any)}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            >
              <option value="string">문자열</option>
              <option value="number">숫자</option>
              <option value="boolean">불린</option>
            </select>
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>설정 값:</label>
            <input
              type={valueType === "number" ? "number" : "text"}
              value={settingValue}
              onChange={(e) => setSettingValue(e.target.value)}
              placeholder={valueType === "boolean" ? "true 또는 false" : valueType === "number" ? "예: 123" : "예: 값"}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={addSetting}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              설정 추가
            </button>
            <button
              onClick={updateSetting}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              값 업데이트
            </button>
          </div>
        </div>

        {/* 설정 조작 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 설정 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={settingKey}
              onChange={(e) => setSettingKey(e.target.value)}
              placeholder="설정 키"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getSetting}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              값 읽기
            </button>
            <button
              onClick={deleteSetting}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              삭제
            </button>
          </div>
        </div>
      </div>

      <div style={{ flex: "1 1 auto", overflowY: "auto", padding: "15px", backgroundColor: "#fff", minHeight: "200px" }}>
        <h4 style={{ marginTop: 0, marginBottom: "10px" }}>결과 및 과정 설명</h4>
        <pre style={{
          backgroundColor: "#f5f5f5",
          padding: "15px",
          borderRadius: "5px",
          whiteSpace: "pre-wrap",
          fontFamily: "monospace",
          fontSize: "12px",
          lineHeight: "1.5",
          margin: 0,
          minHeight: "100px",
        }}>
          {result || "위 버튼을 클릭하여 Excel 설정 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Settings;
