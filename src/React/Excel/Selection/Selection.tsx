import React, { useState } from "react";

/* global Excel */

const Selection: React.FC = () => {
  const [result, setResult] = useState("");
  const [rangeAddress, setRangeAddress] = useState("A1");

  // 1. 현재 선택된 셀/범위 가져오기
  const getCurrentSelection = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address,values,formulas");
        await context.sync();

        const values = range.values as any[][];
        const formulas = range.formulas as string[][];
        
        let resultText = `현재 선택된 범위: ${range.address}\n\n`;
        resultText += `셀 개수: ${values.length}행 × ${values[0]?.length || 0}열\n\n`;
        
        if (values.length === 1 && values[0]?.length === 1) {
          // 단일 셀
          resultText += `값: ${values[0][0] || "(비어있음)"}\n`;
          resultText += `수식: ${formulas[0][0] || "(수식 없음)"}\n`;
        } else {
          // 범위
          resultText += `범위 데이터:\n`;
          values.slice(0, 5).forEach((row, i) => {
            resultText += `  ${row.map(cell => cell || "").join(" | ")}\n`;
          });
          if (values.length > 5) {
            resultText += `  ... (총 ${values.length}행)\n`;
          }
        }

        resultText += `\n과정:\n1. context.workbook.getSelectedRange()으로 현재 선택 범위 가져오기\n2. range.load("address,values,formulas")로 속성 로드\n3. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 특정 셀 선택
  const selectCell = async () => {
    if (!rangeAddress.trim()) {
      setResult("셀 주소를 입력해주세요. (예: A1, B2, C3:D5)");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(rangeAddress);
        range.load("address,values");
        await context.sync();

        range.select();
        await context.sync();

        setResult(`셀 선택 완료!\n주소: ${range.address}\n값: ${JSON.stringify(range.values)}\n\n과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.getRange("${rangeAddress}")로 범위 가져오기\n3. range.select()로 범위 선택\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 활성 셀 가져오기
  const getActiveCell = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address,values,formulas");
        await context.sync();

        // 선택된 범위의 첫 번째 셀을 활성 셀로 간주
        const values = range.values as any[][];
        const formulas = range.formulas as string[][];
        const firstCellValue = values[0]?.[0];
        const firstCellFormula = formulas[0]?.[0];

        setResult(`활성 셀 정보:\n주소: ${range.address}\n값: ${firstCellValue || "(비어있음)"}\n수식: ${firstCellFormula || "(수식 없음)"}\n\n과정:\n1. context.workbook.getSelectedRange()으로 현재 선택된 범위 가져오기\n2. range.load("address,values,formulas")로 속성 로드\n3. context.sync()로 동기화\n4. 선택된 범위의 첫 번째 셀을 활성 셀로 간주`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 전체 시트 선택
  const selectEntireSheet = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load("address,rowCount,columnCount");
        await context.sync();

        if (usedRange) {
          usedRange.select();
          await context.sync();

          setResult(`사용된 범위 선택 완료!\n주소: ${usedRange.address}\n행: ${usedRange.rowCount}, 열: ${usedRange.columnCount}\n\n과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.getUsedRange()로 사용된 범위 가져오기\n3. usedRange.select()로 범위 선택\n4. context.sync()로 동기화`);
        } else {
          setResult("시트에 데이터가 없습니다.");
        }
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 행 선택
  const selectRow = async (rowNumber: number) => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(`${rowNumber}:${rowNumber}`);
        range.load("address");
        await context.sync();

        range.select();
        await context.sync();

        setResult(`행 선택 완료!\n행 번호: ${rowNumber}\n주소: ${range.address}\n\n과정:\n1. sheet.getRange("${rowNumber}:${rowNumber}")로 행 범위 가져오기\n2. range.select()로 행 선택\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. 열 선택
  const selectColumn = async (columnLetter: string) => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(`${columnLetter}:${columnLetter}`);
        range.load("address");
        await context.sync();

        range.select();
        await context.sync();

        setResult(`열 선택 완료!\n열: ${columnLetter}\n주소: ${range.address}\n\n과정:\n1. sheet.getRange("${columnLetter}:${columnLetter}")로 열 범위 가져오기\n2. range.select()로 열 선택\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 셀/범위 선택</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📝 Excel 선택 기능 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel Add-in에서 셀과 범위를 선택하는 다양한 방법을 테스트할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>현재 선택된 범위 가져오기</li>
            <li>특정 셀/범위 선택 (주소로)</li>
            <li>활성 셀 정보 가져오기</li>
            <li>사용된 범위 전체 선택</li>
            <li>행/열 선택</li>
          </ul>
        </div>

        {/* 현재 선택 가져오기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #2196f3" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>🔍 현재 선택 확인</h4>
          <button
            onClick={getCurrentSelection}
            style={{
              padding: "8px 16px",
              backgroundColor: "#2196f3",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              marginRight: "10px",
            }}
          >
            현재 선택 가져오기
          </button>
          <button
            onClick={getActiveCell}
            style={{
              padding: "8px 16px",
              backgroundColor: "#2196f3",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            활성 셀 정보
          </button>
        </div>

        {/* 특정 셀/범위 선택 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📍 셀/범위 선택</h4>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <label style={{ fontSize: "13px" }}>셀 주소:</label>
            <input
              type="text"
              value={rangeAddress}
              onChange={(e) => setRangeAddress(e.target.value)}
              placeholder="예: A1, B2, C3:D5"
              style={{
                flex: 1,
                padding: "8px",
                border: "1px solid #ddd",
                borderRadius: "5px",
              }}
            />
            <button
              onClick={selectCell}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              선택
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666" }}>
            예: A1 (단일 셀), A1:B5 (범위), 1:1 (1행 전체), A:A (A열 전체)
          </div>
        </div>

        {/* 빠른 선택 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>⚡ 빠른 선택</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={() => selectEntireSheet()}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              사용된 범위 전체
            </button>
            <button
              onClick={() => selectRow(1)}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              1행 선택
            </button>
            <button
              onClick={() => selectColumn("A")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              A열 선택
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
          {result || "위 버튼을 클릭하여 Excel 셀 선택 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Selection;
