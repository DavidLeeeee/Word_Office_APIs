import React, { useState } from "react";

/* global Excel */

const Range: React.FC = () => {
  const [result, setResult] = useState("");
  const [rangeAddress, setRangeAddress] = useState("A1");
  const [useSelection, setUseSelection] = useState(false);
  const [cellValue, setCellValue] = useState("");
  const [cellFormula, setCellFormula] = useState("");
  const [rangeValues, setRangeValues] = useState("");
  const [rangeFormulas, setRangeFormulas] = useState("");

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

        setRangeAddress(range.address);
        setUseSelection(true);
        setResult(`선택된 범위를 가져왔습니다!\n주소: ${range.address}\n\n이제 "선택된 범위 사용" 모드가 활성화되었습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 1. 범위 데이터 읽기 (값)
  const readRangeValues = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address,values,text,valueTypes");
        await context.sync();

        const values = range.values as any[][];
        const text = range.text as string[][];
        const valueTypes = range.valueTypes as Excel.RangeValueType[][];

        let resultText = `범위 데이터 읽기 완료!\n주소: ${range.address}\n\n`;
        resultText += `데이터 (${values.length}행 × ${values[0]?.length || 0}열):\n`;

        values.forEach((row, i) => {
          const rowText = text[i]?.map((t, j) => {
            const val = values[i][j];
            const type = valueTypes[i]?.[j];
            return `${t || val || "(비어있음)"} [${type || "Unknown"}]`;
          }).join(" | ") || "";
          resultText += `  ${i + 1}: ${rowText}\n`;
        });

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        resultText += `\n과정:\n1. ${method}로 범위 가져오기\n2. range.load("address,values,text,valueTypes")로 속성 로드\n3. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 범위 수식 읽기
  const readRangeFormulas = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address,formulas");
        await context.sync();

        const formulas = range.formulas as string[][];

        let resultText = `범위 수식 읽기 완료!\n주소: ${range.address}\n\n`;
        resultText += `수식 (${formulas.length}행 × ${formulas[0]?.length || 0}열):\n`;

        formulas.forEach((row, i) => {
          const rowText = row.map(f => f || "(수식 없음)").join(" | ");
          resultText += `  ${i + 1}: ${rowText}\n`;
        });

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        resultText += `\n과정:\n1. ${method}로 범위 가져오기\n2. range.load("address,formulas")로 수식 로드\n3. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 단일 셀 값 쓰기
  const writeCellValue = async () => {
    if (!cellValue.trim()) {
      setResult("값을 입력해주세요.");
      return;
    }
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address");
        await context.sync();

        range.values = [[cellValue]];
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`셀 값 쓰기 완료!\n주소: ${range.address}\n값: ${cellValue}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.values = [["${cellValue}"]]로 값 설정\n3. context.sync()로 동기화`);
        setCellValue("");
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 단일 셀 수식 쓰기
  const writeCellFormula = async () => {
    if (!cellFormula.trim()) {
      setResult("수식을 입력해주세요.");
      return;
    }
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address");
        await context.sync();

        range.formulas = [[cellFormula]];
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`셀 수식 쓰기 완료!\n주소: ${range.address}\n수식: ${cellFormula}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.formulas = [["${cellFormula}"]]로 수식 설정\n3. context.sync()로 동기화`);
        setCellFormula("");
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 범위 값 쓰기 (2차원 배열)
  const writeRangeValues = async () => {
    if (!rangeValues.trim()) {
      setResult("데이터를 입력해주세요.");
      return;
    }
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      // 입력된 텍스트를 2차원 배열로 파싱
      // 예: "1,2,3\n4,5,6" 또는 JSON 형식
      let values: any[][];
      try {
        // JSON 형식 시도
        values = JSON.parse(rangeValues);
        if (!Array.isArray(values) || !Array.isArray(values[0])) {
          throw new Error("Invalid format");
        }
      } catch {
        // 줄바꿈으로 구분된 CSV 형식 파싱
        const lines = rangeValues.trim().split("\n");
        values = lines.map(line => {
          return line.split(",").map(cell => {
            const trimmed = cell.trim();
            // 숫자로 변환 가능하면 숫자로, 아니면 문자열로
            const num = Number(trimmed);
            return isNaN(num) ? trimmed : num;
          });
        });
      }

      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address");
        await context.sync();

        range.values = values;
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`범위 값 쓰기 완료!\n주소: ${range.address}\n데이터: ${values.length}행 × ${values[0]?.length || 0}열\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.values = ${JSON.stringify(values)}로 값 설정\n3. context.sync()로 동기화`);
        setRangeValues("");
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: 데이터는 JSON 배열 형식 (예: [["A","B"],["C","D"]]) 또는 줄바꿈으로 구분된 CSV 형식 (예: A,B\nC,D)으로 입력하세요.`);
    }
  };

  // 6. 범위 수식 쓰기
  const writeRangeFormulas = async () => {
    if (!rangeFormulas.trim()) {
      setResult("수식을 입력해주세요.");
      return;
    }
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      // 입력된 텍스트를 2차원 배열로 파싱
      let formulas: string[][];
      try {
        formulas = JSON.parse(rangeFormulas);
        if (!Array.isArray(formulas) || !Array.isArray(formulas[0])) {
          throw new Error("Invalid format");
        }
      } catch {
        const lines = rangeFormulas.trim().split("\n");
        formulas = lines.map(line => line.split(",").map(f => f.trim()));
      }

      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address");
        await context.sync();

        range.formulas = formulas;
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`범위 수식 쓰기 완료!\n주소: ${range.address}\n수식: ${formulas.length}행 × ${formulas[0]?.length || 0}열\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.formulas = ${JSON.stringify(formulas)}로 수식 설정\n3. context.sync()로 동기화`);
        setRangeFormulas("");
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: 수식은 JSON 배열 형식 또는 줄바꿈으로 구분된 CSV 형식으로 입력하세요.`);
    }
  };

  // 7. 범위 지우기
  const clearRange = async (clearType: "All" | "Formats" | "Contents" | "Hyperlinks" = "All") => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address");
        await context.sync();

        range.clear(clearType);
        await context.sync();

        const clearTypeText = clearType === "All" ? "모두" : clearType === "Formats" ? "서식만" : clearType === "Contents" ? "내용만" : "하이퍼링크만";
        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`범위 지우기 완료!\n주소: ${range.address}\n지우기 유형: ${clearTypeText}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.clear("${clearType}")로 범위 지우기\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 셀 삽입
  const insertCells = async (shift: "Down" | "Right") => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address");
        await context.sync();

        const newRange = range.insert(shift);
        newRange.load("address");
        await context.sync();

        const shiftText = shift === "Down" ? "아래로" : "오른쪽으로";
        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`셀 삽입 완료!\n원본 범위: ${range.address}\n새 범위: ${newRange.address}\n이동 방향: ${shiftText}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.insert("${shift}")로 셀 삽입\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 9. 셀 삭제
  const deleteCells = async (shift: "Up" | "Left") => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        let range: Excel.Range;
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          range = sheet.getRange(rangeAddress);
        }
        range.load("address");
        await context.sync();

        range.delete(shift);
        await context.sync();

        const shiftText = shift === "Up" ? "위로" : "왼쪽으로";
        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`셀 삭제 완료!\n삭제된 범위: ${range.address}\n이동 방향: ${shiftText}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.delete("${shift}")로 셀 삭제\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 10. 범위 복사
  const copyRange = async () => {
    if (!rangeAddress.trim()) {
      setResult("복사할 범위 주소를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sourceRange = context.workbook.getSelectedRange();
        sourceRange.load("address");
        await context.sync();

        if (sourceRange.address === "") {
          setResult("먼저 복사할 범위를 선택해주세요.");
          return;
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const destRange = sheet.getRange(rangeAddress);
        destRange.load("address");
        await context.sync();

        destRange.copyFrom(sourceRange, Excel.RangeCopyType.All);
        await context.sync();

        setResult(`범위 복사 완료!\n원본: ${sourceRange.address}\n대상: ${destRange.address}\n\n과정:\n1. context.workbook.getSelectedRange()으로 선택된 범위 가져오기\n2. sheet.getRange("${rangeAddress}")로 대상 범위 가져오기\n3. destRange.copyFrom(sourceRange, Excel.RangeCopyType.All)로 복사\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: 먼저 Excel에서 복사할 범위를 선택한 후, 대상 범위 주소를 입력하세요.`);
    }
  };

  // 11. 자동 채우기
  const autoFill = async () => {
    if (!rangeAddress.trim()) {
      setResult("자동 채우기할 범위 주소를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sourceRange = context.workbook.getSelectedRange();
        sourceRange.load("address");
        await context.sync();

        if (sourceRange.address === "") {
          setResult("먼저 자동 채우기의 기준이 될 범위를 선택해주세요.");
          return;
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const destRange = sheet.getRange(rangeAddress);
        destRange.load("address");
        await context.sync();

        sourceRange.autoFill(destRange, Excel.AutoFillType.FillDefault);
        await context.sync();

        setResult(`자동 채우기 완료!\n기준 범위: ${sourceRange.address}\n대상 범위: ${destRange.address}\n\n과정:\n1. context.workbook.getSelectedRange()으로 기준 범위 가져오기\n2. sheet.getRange("${rangeAddress}")로 대상 범위 가져오기\n3. sourceRange.autoFill(destRange, Excel.AutoFillType.FillDefault)로 자동 채우기\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: 먼저 Excel에서 자동 채우기의 기준이 될 범위를 선택한 후, 대상 범위 주소를 입력하세요.`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 범위 데이터 입출력</h3>

        {/* 안내 섹션 */}
        <div style={{
          marginBottom: "20px",
          padding: "15px",
          backgroundColor: "#fff3cd",
          borderRadius: "5px",
          border: "1px solid #ffc107",
          fontSize: "13px",
          lineHeight: "1.6"
        }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#856404", fontSize: "14px" }}>📝 범위 데이터 입출력 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#856404" }}>
            Excel 범위의 데이터를 읽고 쓰는 핵심 기능입니다. <strong>가장 중요하고 복잡한</strong> 부분입니다.
          </p>
          <p style={{ margin: "8px 0", color: "#856404", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#856404" }}>
            <li>범위 데이터 읽기 (값, 수식, 텍스트, 타입)</li>
            <li>단일 셀 값/수식 쓰기</li>
            <li>범위 값/수식 쓰기 (2차원 배열)</li>
            <li>범위 지우기 (모두/서식만/내용만)</li>
            <li>셀 삽입/삭제</li>
            <li>범위 복사</li>
            <li>자동 채우기</li>
          </ul>
        </div>

        {/* 범위 주소 입력 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #2196f3" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>📍 범위 지정</h4>
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
                setResult("직접 입력 모드로 전환되었습니다. 범위 주소를 입력하세요.");
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
            value={rangeAddress}
            onChange={(e) => {
              setRangeAddress(e.target.value);
              setUseSelection(false);
            }}
            placeholder={useSelection ? "선택된 범위 사용 중..." : "예: A1, B2, A1:C5"}
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
          {useSelection && (
            <div style={{ fontSize: "12px", color: "#4caf50", marginTop: "5px" }}>
              ✓ Excel에서 범위를 선택한 후 기능을 사용하세요.
            </div>
          )}
        </div>

        {/* 데이터 읽기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📖 데이터 읽기</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={readRangeValues}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              값 읽기
            </button>
            <button
              onClick={readRangeFormulas}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              수식 읽기
            </button>
          </div>
        </div>

        {/* 단일 셀 쓰기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>✍️ 단일 셀 쓰기</h4>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <input
              type="text"
              value={cellValue}
              onChange={(e) => setCellValue(e.target.value)}
              placeholder="셀 값"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={writeCellValue}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              값 쓰기
            </button>
          </div>
          <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
            <input
              type="text"
              value={cellFormula}
              onChange={(e) => setCellFormula(e.target.value)}
              placeholder="셀 수식 (예: =SUM(A1:A5))"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={writeCellFormula}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              수식 쓰기
            </button>
          </div>
        </div>

        {/* 범위 쓰기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>📝 범위 쓰기</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", color: "#666", marginBottom: "5px" }}>
              범위 값 (JSON 배열 또는 CSV 형식):
            </label>
            <textarea
              value={rangeValues}
              onChange={(e) => setRangeValues(e.target.value)}
              placeholder='예: [["A","B"],["C","D"]] 또는 A,B\nC,D'
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minHeight: "60px", fontFamily: "monospace", fontSize: "12px" }}
            />
            <button
              onClick={writeRangeValues}
              style={{
                marginTop: "5px",
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              범위 값 쓰기
            </button>
          </div>
          <div>
            <label style={{ display: "block", fontSize: "12px", color: "#666", marginBottom: "5px" }}>
              범위 수식 (JSON 배열 또는 CSV 형식):
            </label>
            <textarea
              value={rangeFormulas}
              onChange={(e) => setRangeFormulas(e.target.value)}
              placeholder='예: [["=SUM(A1:A5)","=AVERAGE(B1:B5)"]] 또는 =SUM(A1),=AVERAGE(B1)'
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minHeight: "60px", fontFamily: "monospace", fontSize: "12px" }}
            />
            <button
              onClick={writeRangeFormulas}
              style={{
                marginTop: "5px",
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              범위 수식 쓰기
            </button>
          </div>
        </div>

        {/* 범위 조작 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #f44336" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#f44336" }}>🗑️ 범위 지우기</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={() => clearRange("All")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              모두 지우기
            </button>
            <button
              onClick={() => clearRange("Formats")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              서식만 지우기
            </button>
            <button
              onClick={() => clearRange("Contents")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              내용만 지우기
            </button>
          </div>
        </div>

        {/* 셀 삽입/삭제 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>➕➖ 셀 삽입/삭제</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={() => insertCells("Down")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              셀 삽입 (아래로)
            </button>
            <button
              onClick={() => insertCells("Right")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              셀 삽입 (오른쪽으로)
            </button>
            <button
              onClick={() => deleteCells("Up")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              셀 삭제 (위로)
            </button>
            <button
              onClick={() => deleteCells("Left")}
              style={{
                padding: "8px 16px",
                backgroundColor: "#607d8b",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              셀 삭제 (왼쪽으로)
            </button>
          </div>
        </div>

        {/* 복사 및 자동 채우기 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #e91e63" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#e91e63" }}>📋 복사 및 자동 채우기</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={copyRange}
              style={{
                padding: "8px 16px",
                backgroundColor: "#e91e63",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              범위 복사
            </button>
            <button
              onClick={autoFill}
              style={{
                padding: "8px 16px",
                backgroundColor: "#e91e63",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              자동 채우기
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
            참고: 복사/자동 채우기는 먼저 Excel에서 원본 범위를 선택한 후, 대상 범위 주소를 입력하세요.
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
          {result || "위 버튼을 클릭하여 Excel 범위 데이터 입출력 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Range;
