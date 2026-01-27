import React, { useState } from "react";

/* global Excel */

const Audit: React.FC = () => {
  const [result, setResult] = useState("");
  const [rangeAddress, setRangeAddress] = useState("A1");
  const [useSelection, setUseSelection] = useState(false);

  // 범위 가져오기 헬퍼
  const getRange = async (context: Excel.RequestContext, address: string) => {
    if (useSelection) {
      return context.workbook.getSelectedRange();
    } else {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      return sheet.getRange(address);
    }
  };

  // 1. 직접 선행 셀 가져오기
  const getDirectPrecedents = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const precedents = range.getDirectPrecedents();
        precedents.load("address");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`직접 선행 셀 가져오기 완료!\n범위: ${range.address}\n선행 셀: ${precedents.address || "없음"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.getDirectPrecedents()로 직접 선행 셀 가져오기\n3. precedents.load("address")로 주소 로드\n4. context.sync()로 동기화\n\n참고: 선행 셀은 현재 셀의 수식에서 참조하는 직접적인 셀입니다.`);
      });
    } catch (error: any) {
      if (error.code === "ItemNotFound") {
        setResult(`직접 선행 셀을 찾을 수 없습니다.\n범위: ${rangeAddress}\n\n참고: 해당 범위의 셀에 수식이 없거나 선행 셀이 없을 수 있습니다.`);
      } else {
        setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
      }
    }
  };

  // 2. 모든 선행 셀 가져오기
  const getPrecedents = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const precedents = range.getPrecedents();
        precedents.load("address");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`모든 선행 셀 가져오기 완료!\n범위: ${range.address}\n선행 셀: ${precedents.address || "없음"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.getPrecedents()로 모든 선행 셀 가져오기\n3. precedents.load("address")로 주소 로드\n4. context.sync()로 동기화\n\n참고: 모든 선행 셀은 직접 및 간접 선행 셀을 포함합니다.`);
      });
    } catch (error: any) {
      if (error.code === "ItemNotFound") {
        setResult(`선행 셀을 찾을 수 없습니다.\n범위: ${rangeAddress}\n\n참고: 해당 범위의 셀에 수식이 없거나 선행 셀이 없을 수 있습니다.`);
      } else {
        setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
      }
    }
  };

  // 3. 직접 종속 셀 가져오기
  const getDirectDependents = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const dependents = range.getDirectDependents();
        dependents.load("address");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`직접 종속 셀 가져오기 완료!\n범위: ${range.address}\n종속 셀: ${dependents.address || "없음"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.getDirectDependents()로 직접 종속 셀 가져오기\n3. dependents.load("address")로 주소 로드\n4. context.sync()로 동기화\n\n참고: 종속 셀은 현재 셀을 참조하는 셀입니다.`);
      });
    } catch (error: any) {
      if (error.code === "ItemNotFound") {
        setResult(`직접 종속 셀을 찾을 수 없습니다.\n범위: ${rangeAddress}\n\n참고: 해당 범위의 셀을 참조하는 다른 셀이 없을 수 있습니다.`);
      } else {
        setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
      }
    }
  };

  // 4. 모든 종속 셀 가져오기
  const getDependents = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const dependents = range.getDependents();
        dependents.load("address");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`모든 종속 셀 가져오기 완료!\n범위: ${range.address}\n종속 셀: ${dependents.address || "없음"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.getDependents()로 모든 종속 셀 가져오기\n3. dependents.load("address")로 주소 로드\n4. context.sync()로 동기화\n\n참고: 모든 종속 셀은 직접 및 간접 종속 셀을 포함합니다.`);
      });
    } catch (error: any) {
      if (error.code === "ItemNotFound") {
        setResult(`종속 셀을 찾을 수 없습니다.\n범위: ${rangeAddress}\n\n참고: 해당 범위의 셀을 참조하는 다른 셀이 없을 수 있습니다.`);
      } else {
        setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
      }
    }
  };

  // 5. 추적 화살표 지우기
  const clearArrows = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.clearArrows();
        await context.sync();

        setResult(`추적 화살표 지우기 완료!\n\n과정:\n1. context.workbook.worksheets.getActiveWorksheet()로 활성 시트 가져오기\n2. sheet.clearArrows()로 추적 화살표 지우기\n3. context.sync()로 동기화\n\n참고: 워크시트의 모든 추적 화살표가 제거됩니다.`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 선택된 범위 가져오기
  const getSelectedRange = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        setRangeAddress(range.address);
        setUseSelection(true);
        setResult(`선택된 범위: ${range.address}\n이제 '선택된 범위 사용' 모드가 활성화되었습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: Excel에서 범위를 선택한 후 다시 시도해주세요.`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 검사</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>🔍 Excel 검사 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 검사 기능은 수식의 의존성을 추적하는 기능입니다. 선행 셀(수식에서 참조하는 셀)과 종속 셀(현재 셀을 참조하는 셀)을 찾을 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>직접 선행 셀 찾기 (getDirectPrecedents)</li>
            <li>모든 선행 셀 찾기 (getPrecedents)</li>
            <li>직접 종속 셀 찾기 (getDirectDependents)</li>
            <li>모든 종속 셀 찾기 (getDependents)</li>
            <li>추적 화살표 지우기 (clearArrows)</li>
          </ul>
        </div>

        {/* 범위 선택 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📊 범위 선택</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>범위 주소:</label>
            <input
              type="text"
              value={rangeAddress}
              onChange={(e) => setRangeAddress(e.target.value)}
              placeholder="예: A1"
              disabled={useSelection}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px", opacity: useSelection ? 0.6 : 1 }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "flex", alignItems: "center", fontSize: "12px" }}>
              <input
                type="checkbox"
                checked={useSelection}
                onChange={(e) => {
                  setUseSelection(e.target.checked);
                  if (e.target.checked) {
                    getSelectedRange();
                  }
                }}
                style={{ marginRight: "8px" }}
              />
              선택된 범위 사용
            </label>
          </div>
          {!useSelection && (
            <button
              onClick={getSelectedRange}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              현재 선택된 범위 가져오기
            </button>
          )}
        </div>

        {/* 선행 셀 찾기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>⬅️ 선행 셀 찾기</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getDirectPrecedents}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              직접 선행 셀
            </button>
            <button
              onClick={getPrecedents}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              모든 선행 셀
            </button>
          </div>
        </div>

        {/* 종속 셀 찾기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>➡️ 종속 셀 찾기</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getDirectDependents}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              직접 종속 셀
            </button>
            <button
              onClick={getDependents}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              모든 종속 셀
            </button>
          </div>
        </div>

        {/* 추적 화살표 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>🧹 추적 화살표</h4>
          <button
            onClick={clearArrows}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            추적 화살표 지우기
          </button>
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
          {result || "위 버튼을 클릭하여 Excel 검사 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Audit;
