import React, { useState } from "react";

/* global Excel */

const Table: React.FC = () => {
  const [result, setResult] = useState("");
  const [tableName, setTableName] = useState("");
  const [tableAddress, setTableAddress] = useState("A1");
  const [useSelection, setUseSelection] = useState(false);
  const [hasHeaders, setHasHeaders] = useState(true);
  const [newTableName, setNewTableName] = useState("");
  const [tableStyle, setTableStyle] = useState("TableStyleMedium2");

  // 현재 선택된 범위 가져오기
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

        setTableAddress(range.address);
        setUseSelection(true);
        setResult(`선택된 범위를 가져왔습니다!\n주소: ${range.address}\n\n이제 "선택된 범위 사용" 모드가 활성화되었습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 1. 테이블 목록 가져오기
  const listTables = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        tables.load("items/name,items/id,items/showHeaders,items/showTotals,items/rowCount,items/columnCount");
        await context.sync();

        if (tables.items.length === 0) {
          setResult("현재 워크시트에 테이블이 없습니다.");
          return;
        }

        let resultText = `테이블 목록 (${tables.items.length}개):\n\n`;
        tables.items.forEach((table, index) => {
          resultText += `${index + 1}. ${table.name}\n`;
          resultText += `   ID: ${table.id}\n`;
          resultText += `   행: ${table.rowCount}, 열: ${table.columnCount}\n`;
          resultText += `   헤더 표시: ${table.showHeaders ? "예" : "아니오"}\n`;
          resultText += `   합계 행 표시: ${table.showTotals ? "예" : "아니오"}\n\n`;
        });

        resultText += `과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.tables로 테이블 컬렉션 가져오기\n3. tables.load("items/name,items/id,...")로 속성 로드\n4. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 테이블 생성
  const createTable = async () => {
    if (!useSelection && !tableAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let range: Excel.Range;
        
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          range = sheet.getRange(tableAddress);
        }
        
        range.load("address");
        await context.sync();

        const tables = sheet.tables;
        const newTable = tables.add(range, hasHeaders);
        newTable.load("name,id,rowCount,columnCount,showHeaders");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${tableAddress}")`;
        setResult(`테이블 생성 완료!\n범위: ${range.address}\n테이블 이름: ${newTable.name}\nID: ${newTable.id}\n행: ${newTable.rowCount}, 열: ${newTable.columnCount}\n헤더 포함: ${hasHeaders ? "예" : "아니오"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. sheet.tables.add(range, ${hasHeaders})로 테이블 생성\n3. newTable.load()로 속성 로드\n4. context.sync()로 동기화`);
        setTableAddress("A1");
        setUseSelection(false);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 범위가 다른 테이블과 겹치거나 유효하지 않은 경우 생성할 수 없습니다.`);
    }
  };

  // 3. 테이블 정보 읽기
  const getTableInfo = async () => {
    if (!tableName.trim()) {
      setResult("테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        
        table.load("name,id,showHeaders,showTotals,showBandedRows,showBandedColumns,showFilterButton,highlightFirstColumn,highlightLastColumn,style,rowCount,columnCount");
        const range = table.getRange();
        range.load("address");
        await context.sync();

        const info = `테이블 정보:\n\n이름: ${table.name}\nID: ${table.id}\n범위: ${range.address}\n행: ${table.rowCount}, 열: ${table.columnCount}\n헤더 표시: ${table.showHeaders ? "예" : "아니오"}\n합계 행 표시: ${table.showTotals ? "예" : "아니오"}\n줄무늬 행: ${table.showBandedRows ? "예" : "아니오"}\n줄무늬 열: ${table.showBandedColumns ? "예" : "아니오"}\n필터 버튼: ${table.showFilterButton ? "예" : "아니오"}\n첫 열 강조: ${table.highlightFirstColumn ? "예" : "아니오"}\n마지막 열 강조: ${table.highlightLastColumn ? "예" : "아니오"}\n스타일: ${table.style}\n\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.load()로 속성 로드\n3. table.getRange()로 범위 가져오기\n4. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 테이블 이름을 확인해주세요.`);
    }
  };

  // 4. 테이블 이름 변경
  const renameTable = async () => {
    if (!tableName.trim() || !newTableName.trim()) {
      setResult("현재 이름과 새 이름을 모두 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        table.load("name");
        await context.sync();

        const oldName = table.name;
        table.name = newTableName;
        await context.sync();

        setResult(`테이블 이름 변경 완료!\n이전 이름: ${oldName}\n새 이름: ${table.name}\n\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.name = "${newTableName}"로 이름 변경\n3. context.sync()로 동기화`);
        setTableName("");
        setNewTableName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 테이블 이름은 고유해야 하며 Excel의 명명 규칙을 따라야 합니다.`);
    }
  };

  // 5. 테이블 삭제
  const deleteTable = async () => {
    if (!tableName.trim()) {
      setResult("삭제할 테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        table.load("name");
        await context.sync();

        const deletedName = table.name;
        table.delete();
        await context.sync();

        setResult(`테이블 삭제 완료!\n삭제된 테이블: ${deletedName}\n\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.delete()로 테이블 삭제\n3. context.sync()로 동기화`);
        setTableName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 테이블 스타일 적용
  const applyTableStyle = async () => {
    if (!tableName.trim()) {
      setResult("테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        table.load("name,style");
        await context.sync();

        const oldStyle = table.style;
        table.style = tableStyle;
        await context.sync();

        setResult(`테이블 스타일 적용 완료!\n테이블: ${table.name}\n이전 스타일: ${oldStyle}\n새 스타일: ${table.style}\n\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.style = "${tableStyle}"로 스타일 설정\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 7. 테이블 옵션 설정
  const setTableOptions = async () => {
    if (!tableName.trim()) {
      setResult("테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        table.load("name");
        await context.sync();

        // 옵션들을 state에서 가져와서 설정 (추후 구현)
        setResult(`테이블 옵션 설정 기능은 추후 구현 예정입니다.\n\n현재 테이블: ${table.name}`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 8. 테이블을 범위로 변환
  const convertTableToRange = async () => {
    if (!tableName.trim()) {
      setResult("테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        table.load("name");
        await context.sync();

        const convertedName = table.name;
        const range = table.convertToRange();
        range.load("address");
        await context.sync();

        setResult(`테이블을 범위로 변환 완료!\n테이블: ${convertedName}\n변환된 범위: ${range.address}\n\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.convertToRange()로 테이블을 범위로 변환\n3. context.sync()로 동기화\n\n참고: 모든 데이터는 보존되지만 테이블 기능(필터, 정렬 등)은 제거됩니다.`);
        setTableName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 9. 테이블 데이터 읽기
  const readTableData = async () => {
    if (!tableName.trim()) {
      setResult("테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        
        table.load("name,rowCount,columnCount");
        const dataRange = table.getDataBodyRange();
        dataRange.load("address,values");
        const headerRange = table.getHeaderRowRange();
        headerRange.load("values");
        await context.sync();

        const headers = headerRange.values[0] as any[];
        const data = dataRange.values as any[][];

        let resultText = `테이블 데이터 읽기 완료!\n테이블: ${table.name}\n데이터 범위: ${dataRange.address}\n\n헤더:\n${headers.map((h, i) => `  ${i + 1}. ${h || "(비어있음)"}`).join("\n")}\n\n데이터 (${data.length}행):\n`;
        
        data.slice(0, 10).forEach((row, i) => {
          resultText += `  ${i + 1}: ${row.map(cell => cell || "").join(" | ")}\n`;
        });
        
        if (data.length > 10) {
          resultText += `  ... (총 ${data.length}행)\n`;
        }

        resultText += `\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.getHeaderRowRange()로 헤더 범위 가져오기\n3. table.getDataBodyRange()로 데이터 범위 가져오기\n4. range.load("values")로 값 로드\n5. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 10. 필터 지우기
  const clearTableFilters = async () => {
    if (!tableName.trim()) {
      setResult("테이블 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        table.load("name");
        await context.sync();

        table.clearFilters();
        await context.sync();

        setResult(`테이블 필터 지우기 완료!\n테이블: ${table.name}\n\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.clearFilters()로 모든 필터 제거\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 테이블</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📝 Excel 테이블 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 테이블은 구조화된 데이터를 관리하는 강력한 기능입니다. 필터, 정렬, 자동 확장 등의 기능을 제공합니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>테이블 생성 (범위를 테이블로 변환)</li>
            <li>테이블 목록 조회</li>
            <li>테이블 정보 읽기</li>
            <li>테이블 이름 변경</li>
            <li>테이블 삭제</li>
            <li>테이블 스타일 적용</li>
            <li>테이블 데이터 읽기</li>
            <li>필터 지우기</li>
            <li>테이블을 범위로 변환</li>
          </ul>
        </div>

        {/* 테이블 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 테이블 목록</h4>
          <button
            onClick={listTables}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            테이블 목록 보기
          </button>
        </div>

        {/* 테이블 생성 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 테이블 생성</h4>
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
          <input
            type="text"
            value={tableAddress}
            onChange={(e) => {
              setTableAddress(e.target.value);
              setUseSelection(false);
            }}
            placeholder={useSelection ? "선택된 범위 사용 중..." : "예: A1:C10"}
            disabled={useSelection}
            style={{
              width: "100%",
              padding: "8px",
              border: "1px solid #ddd",
              borderRadius: "5px",
              marginBottom: "10px",
              backgroundColor: useSelection ? "#f5f5f5" : "#fff",
              cursor: useSelection ? "not-allowed" : "text",
            }}
          />
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "flex", alignItems: "center", gap: "10px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={hasHeaders}
                onChange={(e) => setHasHeaders(e.target.checked)}
              />
              <span>헤더 행 포함</span>
            </label>
          </div>
          <button
            onClick={createTable}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            테이블 생성
          </button>
        </div>

        {/* 테이블 조작 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 테이블 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={tableName}
              onChange={(e) => setTableName(e.target.value)}
              placeholder="테이블 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={getTableInfo}
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
              onClick={readTableData}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              데이터 읽기
            </button>
            <button
              onClick={clearTableFilters}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              필터 지우기
            </button>
            <button
              onClick={convertTableToRange}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              범위로 변환
            </button>
            <button
              onClick={deleteTable}
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
              value={newTableName}
              onChange={(e) => setNewTableName(e.target.value)}
              placeholder="새 이름"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={renameTable}
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

        {/* 테이블 스타일 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>🎨 테이블 스타일</h4>
          <div style={{ marginBottom: "10px" }}>
            <select
              value={tableStyle}
              onChange={(e) => setTableStyle(e.target.value)}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            >
              <option value="TableStyleLight1">Light 1</option>
              <option value="TableStyleLight2">Light 2</option>
              <option value="TableStyleLight3">Light 3</option>
              <option value="TableStyleLight4">Light 4</option>
              <option value="TableStyleLight5">Light 5</option>
              <option value="TableStyleLight6">Light 6</option>
              <option value="TableStyleLight7">Light 7</option>
              <option value="TableStyleLight8">Light 8</option>
              <option value="TableStyleLight9">Light 9</option>
              <option value="TableStyleLight10">Light 10</option>
              <option value="TableStyleLight11">Light 11</option>
              <option value="TableStyleLight12">Light 12</option>
              <option value="TableStyleLight13">Light 13</option>
              <option value="TableStyleLight14">Light 14</option>
              <option value="TableStyleLight15">Light 15</option>
              <option value="TableStyleLight16">Light 16</option>
              <option value="TableStyleLight17">Light 17</option>
              <option value="TableStyleLight18">Light 18</option>
              <option value="TableStyleLight19">Light 19</option>
              <option value="TableStyleLight20">Light 20</option>
              <option value="TableStyleLight21">Light 21</option>
              <option value="TableStyleMedium1">Medium 1</option>
              <option value="TableStyleMedium2">Medium 2</option>
              <option value="TableStyleMedium3">Medium 3</option>
              <option value="TableStyleMedium4">Medium 4</option>
              <option value="TableStyleMedium5">Medium 5</option>
              <option value="TableStyleMedium6">Medium 6</option>
              <option value="TableStyleMedium7">Medium 7</option>
              <option value="TableStyleMedium8">Medium 8</option>
              <option value="TableStyleMedium9">Medium 9</option>
              <option value="TableStyleMedium10">Medium 10</option>
              <option value="TableStyleMedium11">Medium 11</option>
              <option value="TableStyleMedium12">Medium 12</option>
              <option value="TableStyleMedium13">Medium 13</option>
              <option value="TableStyleMedium14">Medium 14</option>
              <option value="TableStyleMedium15">Medium 15</option>
              <option value="TableStyleMedium16">Medium 16</option>
              <option value="TableStyleMedium17">Medium 17</option>
              <option value="TableStyleMedium18">Medium 18</option>
              <option value="TableStyleMedium19">Medium 19</option>
              <option value="TableStyleMedium20">Medium 20</option>
              <option value="TableStyleMedium21">Medium 21</option>
              <option value="TableStyleMedium22">Medium 22</option>
              <option value="TableStyleMedium23">Medium 23</option>
              <option value="TableStyleMedium24">Medium 24</option>
              <option value="TableStyleMedium25">Medium 25</option>
              <option value="TableStyleMedium26">Medium 26</option>
              <option value="TableStyleMedium27">Medium 27</option>
              <option value="TableStyleMedium28">Medium 28</option>
              <option value="TableStyleDark1">Dark 1</option>
              <option value="TableStyleDark2">Dark 2</option>
              <option value="TableStyleDark3">Dark 3</option>
              <option value="TableStyleDark4">Dark 4</option>
              <option value="TableStyleDark5">Dark 5</option>
              <option value="TableStyleDark6">Dark 6</option>
              <option value="TableStyleDark7">Dark 7</option>
              <option value="TableStyleDark8">Dark 8</option>
              <option value="TableStyleDark9">Dark 9</option>
              <option value="TableStyleDark10">Dark 10</option>
              <option value="TableStyleDark11">Dark 11</option>
            </select>
          </div>
          <button
            onClick={applyTableStyle}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            스타일 적용
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
          {result || "위 버튼을 클릭하여 Excel 테이블 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Table;
