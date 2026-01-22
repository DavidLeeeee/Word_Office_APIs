import React, { useState } from "react";

/* global Excel */

const Pivot: React.FC = () => {
  const [result, setResult] = useState("");
  const [pivotName, setPivotName] = useState("");
  const [sourceAddress, setSourceAddress] = useState("A1");
  const [destinationAddress, setDestinationAddress] = useState("E1");
  const [useSelection, setUseSelection] = useState(false);
  const [newPivotName, setNewPivotName] = useState("");

  // 현재 선택된 범위 가져오기 (소스 데이터용)
  const getSelectedRange = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();

        if (range.address === "") {
          setResult("Excel에서 범위를 선택한 후 다시 시도해주세요.");
          return;
        }

        setSourceAddress(range.address);
        setUseSelection(true);
        setResult(`선택된 범위를 가져왔습니다!\n주소: ${range.address}\n\n이제 "선택된 범위 사용" 모드가 활성화되었습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 1. 피벗 테이블 목록 가져오기
  const listPivotTables = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const pivotTables = sheet.pivotTables;
        pivotTables.load("items/name,items/id");
        await context.sync();

        if (pivotTables.items.length === 0) {
          setResult("현재 워크시트에 피벗 테이블이 없습니다.");
          return;
        }

        let resultText = `피벗 테이블 목록 (${pivotTables.items.length}개):\n\n`;
        pivotTables.items.forEach((pivot, index) => {
          resultText += `${index + 1}. ${pivot.name}\n`;
          resultText += `   ID: ${pivot.id}\n\n`;
        });

        resultText += `과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.pivotTables로 피벗 테이블 컬렉션 가져오기\n3. pivotTables.load("items/name,items/id")로 속성 로드\n4. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 피벗 테이블 생성
  const createPivotTable = async () => {
    if (!pivotName.trim()) {
      setResult("피벗 테이블 이름을 입력해주세요.");
      return;
    }

    if (!useSelection && !sourceAddress.trim()) {
      setResult("소스 데이터 범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    if (!destinationAddress.trim()) {
      setResult("대상 위치 주소를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let sourceRange: Excel.Range;
        
        if (useSelection) {
          sourceRange = context.workbook.getSelectedRange();
        } else {
          sourceRange = sheet.getRange(sourceAddress);
        }
        
        sourceRange.load("address");
        const destinationRange = sheet.getRange(destinationAddress);
        destinationRange.load("address");
        await context.sync();

        const pivotTables = sheet.pivotTables;
        const newPivot = pivotTables.add(pivotName, sourceRange, destinationRange);
        newPivot.load("name,id");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${sourceAddress}")`;
        setResult(`피벗 테이블 생성 완료!\n이름: ${newPivot.name}\nID: ${newPivot.id}\n소스 데이터: ${sourceRange.address}\n대상 위치: ${destinationRange.address}\n\n과정:\n1. ${method}로 소스 데이터 범위 가져오기\n2. sheet.getRange("${destinationAddress}")로 대상 위치 범위 가져오기\n3. sheet.pivotTables.add("${pivotName}", sourceRange, destinationRange)로 피벗 테이블 생성\n4. newPivot.load()로 속성 로드\n5. context.sync()로 동기화\n\n참고: 피벗 테이블을 구성하려면 Excel UI에서 필드를 행/열/값 영역으로 드래그해야 합니다.`);
        setPivotName("");
        setSourceAddress("A1");
        setDestinationAddress("E1");
        setUseSelection(false);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 소스 데이터 범위나 대상 위치가 유효하지 않거나, 피벗 테이블 이름이 중복되는 경우 생성할 수 없습니다.`);
    }
  };

  // 3. 피벗 테이블 정보 읽기
  const getPivotTableInfo = async () => {
    if (!pivotName.trim()) {
      setResult("피벗 테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const pivotTables = sheet.pivotTables;
        const pivot = pivotTables.getItem(pivotName);
        
        pivot.load("name,id,allowMultipleFiltersPerField,enableDataValueEditing,refreshOnOpen,useCustomSortLists");
        const dataSourceString = pivot.getDataSourceString();
        const dataSourceType = pivot.getDataSourceType();
        await context.sync();

        const info = `피벗 테이블 정보:\n\n이름: ${pivot.name}\nID: ${pivot.id}\n데이터 소스: ${dataSourceString.value || "알 수 없음"}\n데이터 소스 타입: ${dataSourceType.value || "알 수 없음"}\n다중 필터 허용: ${pivot.allowMultipleFiltersPerField ? "예" : "아니오"}\n데이터 값 편집 허용: ${pivot.enableDataValueEditing ? "예" : "아니오"}\n열기 시 새로고침: ${pivot.refreshOnOpen ? "예" : "아니오"}\n사용자 지정 정렬 목록 사용: ${pivot.useCustomSortLists ? "예" : "아니오"}\n\n과정:\n1. sheet.pivotTables.getItem("${pivotName}")로 피벗 테이블 가져오기\n2. pivot.load()로 속성 로드\n3. pivot.getDataSourceString(), getDataSourceType()로 데이터 소스 정보 가져오기\n4. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 피벗 테이블 이름을 확인해주세요.`);
    }
  };

  // 4. 피벗 테이블 이름 변경
  const renamePivotTable = async () => {
    if (!pivotName.trim() || !newPivotName.trim()) {
      setResult("현재 이름과 새 이름을 모두 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const pivotTables = sheet.pivotTables;
        const pivot = pivotTables.getItem(pivotName);
        pivot.load("name");
        await context.sync();

        const oldName = pivot.name;
        pivot.name = newPivotName;
        await context.sync();

        setResult(`피벗 테이블 이름 변경 완료!\n이전 이름: ${oldName}\n새 이름: ${pivot.name}\n\n과정:\n1. sheet.pivotTables.getItem("${pivotName}")로 피벗 테이블 가져오기\n2. pivot.name = "${newPivotName}"로 이름 변경\n3. context.sync()로 동기화`);
        setPivotName("");
        setNewPivotName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 피벗 테이블 이름은 고유해야 하며 Excel의 명명 규칙을 따라야 합니다.`);
    }
  };

  // 5. 피벗 테이블 새로고침
  const refreshPivotTable = async () => {
    if (!pivotName.trim()) {
      setResult("피벗 테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const pivotTables = sheet.pivotTables;
        const pivot = pivotTables.getItem(pivotName);
        pivot.load("name");
        await context.sync();

        pivot.refresh();
        await context.sync();

        setResult(`피벗 테이블 새로고침 완료!\n피벗 테이블: ${pivot.name}\n\n과정:\n1. sheet.pivotTables.getItem("${pivotName}")로 피벗 테이블 가져오기\n2. pivot.refresh()로 데이터 새로고침\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 모든 피벗 테이블 새로고침
  const refreshAllPivotTables = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const pivotTables = sheet.pivotTables;
        pivotTables.load("count");
        await context.sync();

        const count = pivotTables.count;
        pivotTables.refreshAll();
        await context.sync();

        setResult(`모든 피벗 테이블 새로고침 완료!\n새로고침된 피벗 테이블 수: ${count}\n\n과정:\n1. sheet.pivotTables로 피벗 테이블 컬렉션 가져오기\n2. pivotTables.refreshAll()로 모든 피벗 테이블 새로고침\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 7. 피벗 테이블 삭제
  const deletePivotTable = async () => {
    if (!pivotName.trim()) {
      setResult("삭제할 피벗 테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const pivotTables = sheet.pivotTables;
        const pivot = pivotTables.getItem(pivotName);
        pivot.load("name");
        await context.sync();

        const deletedName = pivot.name;
        pivot.delete();
        await context.sync();

        setResult(`피벗 테이블 삭제 완료!\n삭제된 피벗 테이블: ${deletedName}\n\n과정:\n1. sheet.pivotTables.getItem("${pivotName}")로 피벗 테이블 가져오기\n2. pivot.delete()로 피벗 테이블 삭제\n3. context.sync()로 동기화`);
        setPivotName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 피벗 테이블</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📊 Excel 피벗 테이블 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 피벗 테이블은 대량의 데이터를 요약하고 분석하는 강력한 기능입니다. 데이터를 다양한 각도에서 집계하고 분석할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>피벗 테이블 생성 (소스 데이터와 대상 위치 지정)</li>
            <li>피벗 테이블 목록 조회</li>
            <li>피벗 테이블 정보 읽기</li>
            <li>피벗 테이블 이름 변경</li>
            <li>피벗 테이블 새로고침 (단일/전체)</li>
            <li>피벗 테이블 삭제</li>
          </ul>
          <p style={{ margin: "8px 0", color: "#1976d2", fontStyle: "italic" }}>
            ⚠️ 참고: 피벗 테이블 생성 후 필드 배치(행/열/값)는 Excel UI에서 직접 수행해야 합니다.
          </p>
        </div>

        {/* 피벗 테이블 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 피벗 테이블 목록</h4>
          <button
            onClick={listPivotTables}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            피벗 테이블 목록 보기
          </button>
        </div>

        {/* 피벗 테이블 생성 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 피벗 테이블 생성</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={pivotName}
              onChange={(e) => setPivotName(e.target.value)}
              placeholder="피벗 테이블 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <button
              onClick={getSelectedRange}
              style={{
                padding: "8px 16px",
                backgroundColor: useSelection ? "#4caf50" : "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontWeight: useSelection ? "bold" : "normal",
              }}
            >
              {useSelection ? "✓ 선택된 범위 사용 중" : "선택된 범위 사용"}
            </button>
            <button
              onClick={() => {
                setUseSelection(false);
                setResult("직접 입력 모드로 전환되었습니다.");
              }}
              style={{
                padding: "8px 16px",
                backgroundColor: !useSelection ? "#4caf50" : "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontWeight: !useSelection ? "bold" : "normal",
              }}
            >
              직접 입력
            </button>
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>소스 데이터 범위:</label>
            <input
              type="text"
              value={sourceAddress}
              onChange={(e) => {
                setSourceAddress(e.target.value);
                setUseSelection(false);
              }}
              placeholder={useSelection ? "선택된 범위 사용 중..." : "예: A1:D10"}
              disabled={useSelection}
              style={{
                width: "100%",
                padding: "8px",
                border: "1px solid #ddd",
                borderRadius: "5px",
                backgroundColor: useSelection ? "#f5f5f5" : "#fff",
                cursor: useSelection ? "not-allowed" : "text",
              }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>대상 위치 (피벗 테이블이 생성될 위치):</label>
            <input
              type="text"
              value={destinationAddress}
              onChange={(e) => setDestinationAddress(e.target.value)}
              placeholder="예: E1"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
          </div>
          <button
            onClick={createPivotTable}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            피벗 테이블 생성
          </button>
        </div>

        {/* 피벗 테이블 조작 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 피벗 테이블 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={pivotName}
              onChange={(e) => setPivotName(e.target.value)}
              placeholder="피벗 테이블 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={getPivotTableInfo}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              정보 읽기
            </button>
            <button
              onClick={refreshPivotTable}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              새로고침
            </button>
            <button
              onClick={deletePivotTable}
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
          <div style={{ display: "flex", gap: "10px", alignItems: "center", marginBottom: "10px" }}>
            <input
              type="text"
              value={newPivotName}
              onChange={(e) => setNewPivotName(e.target.value)}
              placeholder="새 이름"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={renamePivotTable}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              이름 변경
            </button>
          </div>
        </div>

        {/* 전체 새로고침 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>🔄 전체 새로고침</h4>
          <button
            onClick={refreshAllPivotTables}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            모든 피벗 테이블 새로고침
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
          {result || "위 버튼을 클릭하여 Excel 피벗 테이블 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Pivot;
