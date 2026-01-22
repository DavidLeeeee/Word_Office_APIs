import React, { useState } from "react";

/* global Excel */

const Worksheet: React.FC = () => {
  const [result, setResult] = useState("");
  const [worksheetName, setWorksheetName] = useState("");
  const [newWorksheetName, setNewWorksheetName] = useState("");
  const [newName, setNewName] = useState("");
  const [position, setPosition] = useState("0");
  const [tabColorValue, setTabColorValue] = useState("#FFA500");

  // 1. 활성 워크시트 정보 가져오기
  const getActiveWorksheetInfo = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name,id,position,visibility,tabColor,showGridlines,showHeadings,enableCalculation");
        await context.sync();

        const visibility = sheet.visibility === "Visible" ? "보임" : sheet.visibility === "Hidden" ? "숨김" : "매우 숨김";
        const info = `📋 활성 워크시트 정보:\n\n이름: ${sheet.name}\nID: ${sheet.id}\n위치: ${sheet.position}\n표시 상태: ${visibility}\n탭 색상: ${sheet.tabColor || "(자동)"}\n격자선 표시: ${sheet.showGridlines ? "예" : "아니오"}\n행/열 머리글 표시: ${sheet.showHeadings ? "예" : "아니오"}\n계산 활성화: ${sheet.enableCalculation ? "예" : "아니오"}\n\n과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.load()로 필요한 속성 로드\n3. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 워크시트 추가
  const addWorksheet = async () => {
    const name = newWorksheetName.trim() || undefined;
    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const newSheet = worksheets.add(name);
        newSheet.load("name,id,position");
        await context.sync();

        setResult(`워크시트 추가 완료!\n이름: ${newSheet.name}\nID: ${newSheet.id}\n위치: ${newSheet.position}\n\n과정:\n1. context.workbook.worksheets.add("${name || "(이름 없음)"}")로 워크시트 추가\n2. newSheet.load()로 속성 로드\n3. context.sync()로 동기화`);
        setNewWorksheetName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 워크시트 이름은 고유해야 하며 32자 미만이어야 합니다.`);
    }
  };

  // 3. 워크시트 삭제
  const deleteWorksheet = async () => {
    if (!worksheetName.trim()) {
      setResult("삭제할 워크시트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const sheet = worksheets.getItem(worksheetName);
        sheet.load("name");
        await context.sync();

        sheet.delete();
        await context.sync();

        setResult(`워크시트 삭제 완료!\n삭제된 워크시트: ${sheet.name}\n\n과정:\n1. context.workbook.worksheets.getItem("${worksheetName}")로 워크시트 가져오기\n2. sheet.delete()로 워크시트 삭제\n3. context.sync()로 동기화\n\n⚠️ 참고: "매우 숨김" 상태의 워크시트는 먼저 표시 상태를 변경해야 삭제할 수 있습니다.`);
        setWorksheetName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 워크시트를 찾을 수 없거나, "매우 숨김" 상태인 경우 삭제할 수 없습니다.`);
    }
  };

  // 4. 워크시트 이름 변경
  const renameWorksheet = async () => {
    if (!worksheetName.trim() || !newName.trim()) {
      setResult("현재 이름과 새 이름을 모두 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const sheet = worksheets.getItem(worksheetName);
        sheet.load("name");
        await context.sync();

        const oldName = sheet.name;
        sheet.name = newName;
        await context.sync();

        setResult(`워크시트 이름 변경 완료!\n이전 이름: ${oldName}\n새 이름: ${sheet.name}\n\n과정:\n1. context.workbook.worksheets.getItem("${worksheetName}")로 워크시트 가져오기\n2. sheet.name = "${newName}"로 이름 변경\n3. context.sync()로 동기화`);
        setWorksheetName("");
        setNewName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 워크시트 이름은 고유해야 하며 32자 미만이어야 합니다.`);
    }
  };

  // 5. 워크시트 복사
  const copyWorksheet = async () => {
    if (!worksheetName.trim()) {
      setResult("복사할 워크시트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const sourceSheet = worksheets.getItem(worksheetName);
        sourceSheet.load("name");
        await context.sync();

        const copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.End);
        copiedSheet.load("name,position");
        await context.sync();

        setResult(`워크시트 복사 완료!\n원본: ${sourceSheet.name}\n복사본: ${copiedSheet.name}\n위치: ${copiedSheet.position}\n\n과정:\n1. context.workbook.worksheets.getItem("${worksheetName}")로 워크시트 가져오기\n2. sourceSheet.copy(Excel.WorksheetPositionType.End)로 워크시트 복사\n3. context.sync()로 동기화`);
        setWorksheetName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 워크시트 활성화
  const activateWorksheet = async () => {
    if (!worksheetName.trim()) {
      setResult("활성화할 워크시트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const sheet = worksheets.getItem(worksheetName);
        sheet.load("name");
        await context.sync();

        sheet.activate();
        await context.sync();

        setResult(`워크시트 활성화 완료!\n활성화된 워크시트: ${sheet.name}\n\n과정:\n1. context.workbook.worksheets.getItem("${worksheetName}")로 워크시트 가져오기\n2. sheet.activate()로 워크시트 활성화\n3. context.sync()로 동기화`);
        setWorksheetName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 7. 워크시트 표시/숨김
  const setWorksheetVisibility = async (visibility: "Visible" | "Hidden" | "VeryHidden") => {
    if (!worksheetName.trim()) {
      setResult("워크시트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const sheet = worksheets.getItem(worksheetName);
        sheet.load("name,visibility");
        await context.sync();

        const oldVisibility = sheet.visibility;
        sheet.visibility = visibility;
        await context.sync();

        const visibilityText = visibility === "Visible" ? "보임" : visibility === "Hidden" ? "숨김" : "매우 숨김";
        const oldVisibilityText = oldVisibility === "Visible" ? "보임" : oldVisibility === "Hidden" ? "숨김" : "매우 숨김";

        setResult(`워크시트 표시 상태 변경 완료!\n워크시트: ${sheet.name}\n이전 상태: ${oldVisibilityText}\n새 상태: ${visibilityText}\n\n과정:\n1. context.workbook.worksheets.getItem("${worksheetName}")로 워크시트 가져오기\n2. sheet.visibility = "${visibility}"로 표시 상태 변경\n3. context.sync()로 동기화`);
        setWorksheetName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 8. 워크시트 위치 변경
  const moveWorksheet = async () => {
    if (!worksheetName.trim()) {
      setResult("이동할 워크시트 이름을 입력해주세요.");
      return;
    }

    const pos = parseInt(position);
    if (isNaN(pos) || pos < 0) {
      setResult("올바른 위치 값을 입력해주세요. (0 이상의 정수)");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const sheet = worksheets.getItem(worksheetName);
        sheet.load("name,position");
        await context.sync();

        const oldPosition = sheet.position;
        sheet.position = pos;
        await context.sync();

        setResult(`워크시트 위치 변경 완료!\n워크시트: ${sheet.name}\n이전 위치: ${oldPosition}\n새 위치: ${sheet.position}\n\n과정:\n1. context.workbook.worksheets.getItem("${worksheetName}")로 워크시트 가져오기\n2. sheet.position = ${pos}로 위치 변경\n3. context.sync()로 동기화`);
        setWorksheetName("");
        setPosition("0");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 9. 탭 색상 변경
  const changeTabColor = async () => {
    if (!worksheetName.trim()) {
      setResult("워크시트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        const sheet = worksheets.getItem(worksheetName);
        sheet.load("name,tabColor");
        await context.sync();

        const oldColor = sheet.tabColor || "(자동)";
        sheet.tabColor = tabColorValue;
        await context.sync();

        setResult(`탭 색상 변경 완료!\n워크시트: ${sheet.name}\n이전 색상: ${oldColor}\n새 색상: ${sheet.tabColor}\n\n과정:\n1. context.workbook.worksheets.getItem("${worksheetName}")로 워크시트 가져오기\n2. sheet.tabColor = "${tabColorValue}"로 탭 색상 변경\n3. context.sync()로 동기화`);
        setWorksheetName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 10. 격자선/머리글 표시 설정
  const setDisplayOptions = async (option: "gridlines" | "headings", value: boolean) => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();

        if (option === "gridlines") {
          sheet.showGridlines = value;
        } else {
          sheet.showHeadings = value;
        }
        await context.sync();

        const optionName = option === "gridlines" ? "격자선" : "행/열 머리글";
        setResult(`${optionName} 표시 설정 완료!\n워크시트: ${sheet.name}\n설정: ${value ? "표시" : "숨김"}\n\n과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.show${option === "gridlines" ? "Gridlines" : "Headings"} = ${value}로 설정\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 11. 워크시트 보호 상태 확인
  const checkProtection = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const protection = sheet.protection;
        protection.load("protected");
        await context.sync();

        setResult(`워크시트 보호 상태:\n\n보호됨: ${protection.protected ? "예" : "아니오"}\n\n과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.protection.load("protected")로 보호 상태 로드\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 워크시트</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📝 워크시트 기능 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 워크시트의 생성, 삭제, 이름 변경, 복사, 활성화, 표시/숨김 등을 관리할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>워크시트 정보 확인</li>
            <li>워크시트 추가/삭제</li>
            <li>워크시트 이름 변경</li>
            <li>워크시트 복사</li>
            <li>워크시트 활성화</li>
            <li>워크시트 표시/숨김</li>
            <li>워크시트 위치 변경</li>
            <li>탭 색상 변경</li>
            <li>격자선/머리글 표시 설정</li>
            <li>워크시트 보호 상태 확인</li>
          </ul>
        </div>

        {/* 활성 워크시트 정보 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #2196f3" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>🔍 워크시트 정보</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getActiveWorksheetInfo}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              활성 시트 정보
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

        {/* 워크시트 추가 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>➕ 워크시트 추가</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
            <input
              type="text"
              value={newWorksheetName}
              onChange={(e) => setNewWorksheetName(e.target.value)}
              placeholder="워크시트 이름 (선택사항)"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={addWorksheet}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
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

        {/* 워크시트 조작 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>🔧 워크시트 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={worksheetName}
              onChange={(e) => setWorksheetName(e.target.value)}
              placeholder="워크시트 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={activateWorksheet}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              활성화
            </button>
            <button
              onClick={copyWorksheet}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              복사
            </button>
            <button
              onClick={deleteWorksheet}
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
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <input
              type="text"
              value={newName}
              onChange={(e) => setNewName(e.target.value)}
              placeholder="새 이름"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={renameWorksheet}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              이름 변경
            </button>
          </div>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <input
              type="number"
              value={position}
              onChange={(e) => setPosition(e.target.value)}
              placeholder="위치 (0부터)"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={moveWorksheet}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              위치 변경
            </button>
          </div>
        </div>

        {/* 표시/숨김 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>👁️ 표시/숨김</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={() => setWorksheetVisibility("Visible")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              보이기
            </button>
            <button
              onClick={() => setWorksheetVisibility("Hidden")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              숨기기
            </button>
            <button
              onClick={() => setWorksheetVisibility("VeryHidden")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              매우 숨김
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
            참고: "매우 숨김" 상태는 Excel UI에서 직접 표시할 수 없으며, API로만 접근 가능합니다.
          </div>
        </div>

        {/* 탭 색상 및 표시 옵션 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>🎨 탭 색상 및 표시 옵션</h4>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <label style={{ fontSize: "13px" }}>탭 색상:</label>
            <input
              type="color"
              value={tabColorValue}
              onChange={(e) => setTabColorValue(e.target.value)}
              style={{ padding: "4px", border: "1px solid #ddd", borderRadius: "5px", height: "40px" }}
            />
            <input
              type="text"
              value={worksheetName}
              onChange={(e) => setWorksheetName(e.target.value)}
              placeholder="워크시트 이름"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={changeTabColor}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              색상 적용
            </button>
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={() => setDisplayOptions("gridlines", true)}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              격자선 표시
            </button>
            <button
              onClick={() => setDisplayOptions("gridlines", false)}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              격자선 숨김
            </button>
            <button
              onClick={() => setDisplayOptions("headings", true)}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              머리글 표시
            </button>
            <button
              onClick={() => setDisplayOptions("headings", false)}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              머리글 숨김
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
            참고: 격자선/머리글 표시 설정은 활성 워크시트에 적용됩니다.
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
          {result || "위 버튼을 클릭하여 Excel 워크시트 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Worksheet;
