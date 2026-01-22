import React, { useState } from "react";

/* global Excel */

const Workbook: React.FC = () => {
  const [result, setResult] = useState("");
  const [propName, setPropName] = useState("title");
  const [propValue, setPropValue] = useState("");
  const [customPropName, setCustomPropName] = useState("");
  const [customPropValue, setCustomPropValue] = useState("");

  // 1. 워크북 기본 정보 가져오기
  const getWorkbookInfo = async () => {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name,isDirty,readOnly,autoSave,previouslySaved");
        await context.sync();

        const info = `📋 워크북 기본 정보:\n\n이름: ${workbook.name}\n변경 여부: ${workbook.isDirty ? "변경됨" : "변경 없음"}\n읽기 전용: ${workbook.readOnly ? "예" : "아니오"}\n자동 저장: ${workbook.autoSave ? "활성화" : "비활성화"}\n이전 저장 여부: ${workbook.previouslySaved ? "저장됨" : "저장 안 됨"}\n\n과정:\n1. context.workbook으로 워크북 객체 가져오기\n2. workbook.load("name,isDirty,readOnly,autoSave,previouslySaved")로 속성 로드\n3. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 워크시트 목록 가져오기
  const getWorksheets = async () => {
    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name,items/visibility,items/position");
        await context.sync();

        let info = `📊 워크시트 목록 (${worksheets.items.length}개):\n\n`;
        worksheets.items.forEach((sheet, index) => {
          const visibility = sheet.visibility === "Visible" ? "보임" : sheet.visibility === "Hidden" ? "숨김" : "매우 숨김";
          info += `${index + 1}. ${sheet.name} (위치: ${sheet.position}, 상태: ${visibility})\n`;
        });

        info += `\n과정:\n1. context.workbook.worksheets로 워크시트 컬렉션 가져오기\n2. worksheets.load("items/name,items/visibility,items/position")로 속성 로드\n3. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 워크북 속성 읽기
  const readWorkbookProperties = async () => {
    try {
      await Excel.run(async (context) => {
        const properties = context.workbook.properties;
        properties.load("title,author,subject,keywords,category,comments,company,manager");
        await context.sync();

        const info = `📝 워크북 속성:\n\n제목: ${properties.title || "(없음)"}\n작성자: ${properties.author || "(없음)"}\n주제: ${properties.subject || "(없음)"}\n키워드: ${properties.keywords || "(없음)"}\n카테고리: ${properties.category || "(없음)"}\n설명: ${properties.comments || "(없음)"}\n회사: ${properties.company || "(없음)"}\n관리자: ${properties.manager || "(없음)"}\n\n과정:\n1. context.workbook.properties로 워크북 속성 가져오기\n2. properties.load()로 필요한 속성 로드\n3. context.sync()로 동기화\n\n확인 방법:\n- Excel UI: 파일 > 정보 > 속성 > 고급 속성\n- 또는 파일 탐색기에서 파일 우클릭 > 속성 > 세부 정보`;

        setResult(info);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 워크북 속성 설정
  const setWorkbookProperty = async () => {
    if (!propValue.trim()) {
      setResult("속성 값을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const properties = context.workbook.properties;
        properties.load(propName);
        await context.sync();

        // 각 속성별로 직접 설정
        if (propName === "title") {
          properties.title = propValue;
        } else if (propName === "author") {
          properties.author = propValue;
        } else if (propName === "subject") {
          properties.subject = propValue;
        } else if (propName === "keywords") {
          properties.keywords = propValue;
        } else if (propName === "category") {
          properties.category = propValue;
        } else if (propName === "comments") {
          properties.comments = propValue;
        } else if (propName === "company") {
          properties.company = propValue;
        } else if (propName === "manager") {
          properties.manager = propValue;
        } else {
          setResult(`지원하지 않는 속성입니다: ${propName}`);
          return;
        }

        await context.sync();

        setResult(`${propName} 설정 완료!\n값: ${propValue}\n\n과정:\n1. context.workbook.properties로 워크북 속성 가져오기\n2. properties.load("${propName}")로 속성 로드\n3. properties.${propName} = "${propValue}"로 속성 설정\n4. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 일부 속성은 읽기 전용이거나 현재 Excel 버전/환경에서 지원되지 않을 수 있습니다.`);
    }
  };

  // 5. 커스텀 속성 추가
  const addCustomProperty = async () => {
    if (!customPropName.trim() || !customPropValue.trim()) {
      setResult("커스텀 속성 이름과 값을 모두 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const properties = context.workbook.properties;
        const customProps = properties.custom;
        customProps.add(customPropName, customPropValue);
        await context.sync();

        setResult(`커스텀 속성 추가 완료!\n이름: ${customPropName}\n값: ${customPropValue}\n\n과정:\n1. context.workbook.properties.custom으로 커스텀 속성 컬렉션 가져오기\n2. customProps.add("${customPropName}", "${customPropValue}")로 속성 추가\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 커스텀 속성 목록
  const listCustomProperties = async () => {
    try {
      await Excel.run(async (context) => {
        const properties = context.workbook.properties;
        const customProps = properties.custom;
        customProps.load("items/key,items/value");
        await context.sync();

        if (customProps.items.length === 0) {
          setResult("커스텀 속성이 없습니다.");
          return;
        }

        const propList = customProps.items.map((prop: any, idx: number) => {
          return `${idx + 1}. ${prop.key}: ${prop.value}`;
        }).join("\n");

        setResult(`커스텀 속성 목록 (${customProps.items.length}개):\n\n${propList}\n\n과정:\n1. context.workbook.properties.custom으로 커스텀 속성 컬렉션 가져오기\n2. customProps.load("items/key,items/value")로 속성 로드\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 7. 워크북 보호 상태 확인
  const checkProtection = async () => {
    try {
      await Excel.run(async (context) => {
        const protection = context.workbook.protection;
        protection.load("protected");
        await context.sync();

        setResult(`워크북 보호 상태:\n\n보호됨: ${protection.protected ? "예" : "아니오"}\n\n과정:\n1. context.workbook.protection으로 보호 객체 가져오기\n2. protection.load("protected")로 속성 로드\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 워크북 설정 읽기
  const readSettings = async () => {
    try {
      await Excel.run(async (context) => {
        const settings = context.workbook.settings;
        settings.load("items");
        await context.sync();

        if (settings.items.length === 0) {
          setResult("워크북 설정이 없습니다.");
          return;
        }

        const settingList = settings.items.map((setting: any, idx: number) => {
          return `${idx + 1}. ${setting.key}: ${setting.value}`;
        }).join("\n");

        setResult(`워크북 설정 목록 (${settings.items.length}개):\n\n${settingList}\n\n과정:\n1. context.workbook.settings으로 설정 컬렉션 가져오기\n2. settings.load("items")로 설정 로드\n3. context.sync()로 동기화\n\n참고: 설정은 이 Add-in 전용으로 워크북에 저장됩니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 워크북</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📝 워크북 기능 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 워크북의 기본 정보, 속성, 설정 등을 관리할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>워크북 기본 정보 (이름, 변경 여부, 읽기 전용 등)</li>
            <li>워크시트 목록</li>
            <li>워크북 속성 (제목, 작성자, 주제 등)</li>
            <li>커스텀 속성</li>
            <li>워크북 보호 상태</li>
            <li>워크북 설정 (Add-in 전용)</li>
          </ul>
        </div>

        {/* 워크북 정보 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #2196f3" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>🔍 워크북 정보</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getWorkbookInfo}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              기본 정보
            </button>
            <button
              onClick={getWorksheets}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              워크시트 목록
            </button>
            <button
              onClick={checkProtection}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              보호 상태
            </button>
          </div>
        </div>

        {/* 워크북 속성 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📝 워크북 속성</h4>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px" }}>
            <button
              onClick={readWorkbookProperties}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              속성 읽기
            </button>
          </div>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <select
              value={propName}
              onChange={(e) => setPropName(e.target.value)}
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            >
              <option value="title">제목</option>
              <option value="author">작성자</option>
              <option value="subject">주제</option>
              <option value="keywords">키워드</option>
              <option value="category">카테고리</option>
              <option value="comments">설명</option>
              <option value="company">회사</option>
              <option value="manager">관리자</option>
            </select>
            <input
              type="text"
              value={propValue}
              onChange={(e) => setPropValue(e.target.value)}
              placeholder="속성 값 입력"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={setWorkbookProperty}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              설정
            </button>
          </div>
        </div>

        {/* 커스텀 속성 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>🔧 커스텀 속성</h4>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px" }}>
            <button
              onClick={listCustomProperties}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              목록 보기
            </button>
          </div>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px" }}>
            <input
              type="text"
              value={customPropName}
              onChange={(e) => setCustomPropName(e.target.value)}
              placeholder="속성 이름"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <input
              type="text"
              value={customPropValue}
              onChange={(e) => setCustomPropValue(e.target.value)}
              placeholder="속성 값"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={addCustomProperty}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              추가
            </button>
          </div>
        </div>

        {/* 워크북 설정 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>⚙️ 워크북 설정</h4>
          <button
            onClick={readSettings}
            style={{
              padding: "8px 16px",
              backgroundColor: "#9c27b0",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            설정 읽기
          </button>
          <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
            참고: 설정은 이 Add-in 전용으로 워크북에 저장됩니다.
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
          {result || "위 버튼을 클릭하여 Excel 워크북 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Workbook;
