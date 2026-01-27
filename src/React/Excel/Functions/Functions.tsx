import React, { useState } from "react";

/* global Excel */

const Functions: React.FC = () => {
  const [result, setResult] = useState("");
  const [functionName, setFunctionName] = useState("SUM");
  const [functionArgs, setFunctionArgs] = useState("1,2,3");
  const [rangeAddress, setRangeAddress] = useState("A1:A5");
  const [useSelection, setUseSelection] = useState(false);
  
  // 수식 입력용 상태
  const [formulaTargetRange, setFormulaTargetRange] = useState("C1:C10");
  const [formulaExpression, setFormulaExpression] = useState("=A1+B1");
  const [useSelectionForFormula, setUseSelectionForFormula] = useState(false);

  // 범위 가져오기 헬퍼
  const getRange = async (context: Excel.RequestContext, address: string) => {
    if (useSelection) {
      return context.workbook.getSelectedRange();
    } else {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      return sheet.getRange(address);
    }
  };

  // 1. SUM 함수
  const calculateSum = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const functions = context.workbook.functions;
        const sumResult = functions.sum(range);
        sumResult.load("value,error");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        if (sumResult.error) {
          setResult(`SUM 함수 계산 오류!\n범위: ${range.address}\n오류: ${sumResult.error}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.sum(range)로 계산\n3. sumResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        } else {
          setResult(`SUM 함수 계산 완료!\n범위: ${range.address}\n결과: ${sumResult.value}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.sum(range)로 계산\n3. sumResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        }
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 2. AVERAGE 함수
  const calculateAverage = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const functions = context.workbook.functions;
        const avgResult = functions.average(range);
        avgResult.load("value,error");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        if (avgResult.error) {
          setResult(`AVERAGE 함수 계산 오류!\n범위: ${range.address}\n오류: ${avgResult.error}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.average(range)로 계산\n3. avgResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        } else {
          setResult(`AVERAGE 함수 계산 완료!\n범위: ${range.address}\n결과: ${avgResult.value}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.average(range)로 계산\n3. avgResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        }
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 3. COUNT 함수
  const calculateCount = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const functions = context.workbook.functions;
        const countResult = functions.count(range);
        countResult.load("value,error");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        if (countResult.error) {
          setResult(`COUNT 함수 계산 오류!\n범위: ${range.address}\n오류: ${countResult.error}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.count(range)로 계산\n3. countResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        } else {
          setResult(`COUNT 함수 계산 완료!\n범위: ${range.address}\n결과: ${countResult.value}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.count(range)로 계산\n3. countResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        }
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 4. MAX 함수
  const calculateMax = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const functions = context.workbook.functions;
        const maxResult = functions.max(range);
        maxResult.load("value,error");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        if (maxResult.error) {
          setResult(`MAX 함수 계산 오류!\n범위: ${range.address}\n오류: ${maxResult.error}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.max(range)로 계산\n3. maxResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        } else {
          setResult(`MAX 함수 계산 완료!\n범위: ${range.address}\n결과: ${maxResult.value}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.max(range)로 계산\n3. maxResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        }
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 5. MIN 함수
  const calculateMin = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const range = await getRange(context, rangeAddress);
        range.load("address");
        await context.sync();

        const functions = context.workbook.functions;
        const minResult = functions.min(range);
        minResult.load("value,error");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        if (minResult.error) {
          setResult(`MIN 함수 계산 오류!\n범위: ${range.address}\n오류: ${minResult.error}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.min(range)로 계산\n3. minResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        } else {
          setResult(`MIN 함수 계산 완료!\n범위: ${range.address}\n결과: ${minResult.value}\n\n과정:\n1. ${method}로 범위 가져오기\n2. context.workbook.functions.min(range)로 계산\n3. minResult.load("value,error")로 결과 로드\n4. context.sync()로 동기화`);
        }
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. ABS 함수 (단일 값)
  const calculateAbs = async () => {
    if (!functionArgs.trim()) {
      setResult("숫자를 입력해주세요.");
      return;
    }

    try {
      const num = parseFloat(functionArgs);
      if (isNaN(num)) {
        setResult("유효한 숫자를 입력해주세요.");
        return;
      }

      await Excel.run(async (context) => {
        const functions = context.workbook.functions;
        const absResult = functions.abs(num);
        absResult.load("value,error");
        await context.sync();

        if (absResult.error) {
          setResult(`ABS 함수 계산 오류!\n입력: ${num}\n오류: ${absResult.error}\n\n과정:\n1. context.workbook.functions.abs(${num})로 계산\n2. absResult.load("value,error")로 결과 로드\n3. context.sync()로 동기화`);
        } else {
          setResult(`ABS 함수 계산 완료!\n입력: ${num}\n결과: ${absResult.value}\n\n과정:\n1. context.workbook.functions.abs(${num})로 계산\n2. absResult.load("value,error")로 결과 로드\n3. context.sync()로 동기화`);
        }
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

  // 수식 입력용 범위 가져오기
  const getSelectedRangeForFormula = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        setFormulaTargetRange(range.address);
        setUseSelectionForFormula(true);
        setResult(`선택된 범위: ${range.address}\n이제 '선택된 범위 사용' 모드가 활성화되었습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: Excel에서 범위를 선택한 후 다시 시도해주세요.`);
    }
  };

  // 수식을 셀에 넣기
  const insertFormula = async () => {
    if (!useSelectionForFormula && !formulaTargetRange.trim()) {
      setResult("대상 범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }
    if (!formulaExpression.trim()) {
      setResult("수식을 입력해주세요. (예: =A1+B1)");
      return;
    }

    try {
      // 수식이 =로 시작하지 않으면 자동으로 추가
      let formula = formulaExpression.trim();
      if (!formula.startsWith("=")) {
        formula = "=" + formula;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let targetRange: Excel.Range;
        
        if (useSelectionForFormula) {
          targetRange = context.workbook.getSelectedRange();
        } else {
          targetRange = sheet.getRange(formulaTargetRange);
        }
        
        targetRange.load("address,rowCount,columnCount");
        await context.sync();

        // 범위의 각 셀에 수식 적용
        // Excel은 자동으로 상대 참조를 처리합니다
        const rowCount = targetRange.rowCount;
        const colCount = targetRange.columnCount;
        
        // 2D 배열로 수식 생성 (각 셀에 동일한 수식, Excel이 자동으로 상대 참조 처리)
        const formulas: string[][] = [];
        for (let i = 0; i < rowCount; i++) {
          const row: string[] = [];
          for (let j = 0; j < colCount; j++) {
            row.push(formula);
          }
          formulas.push(row);
        }

        targetRange.formulas = formulas;
        await context.sync();

        const method = useSelectionForFormula ? "context.workbook.getSelectedRange()" : `sheet.getRange("${formulaTargetRange}")`;
        setResult(`수식 입력 완료!\n대상 범위: ${targetRange.address}\n적용된 수식: ${formula}\n셀 개수: ${rowCount}행 × ${colCount}열\n\n과정:\n1. ${method}로 대상 범위 가져오기\n2. range.formulas = [수식 배열]로 수식 설정\n3. context.sync()로 동기화\n\n참고: Excel이 자동으로 각 셀에 상대 참조를 적용합니다.\n예: C1에 =A1+B1을 넣으면, C2에는 =A2+B2가 자동으로 적용됩니다.`);
        setFormulaExpression("=A1+B1");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 수식 형식이 올바른지 확인해주세요.`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 함수</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>🔢 Excel 함수 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 함수는 워크북의 내장 함수를 JavaScript에서 호출하여 계산할 수 있는 기능입니다. 다양한 수학, 통계, 논리 함수를 사용할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>SUM, AVERAGE, COUNT, MAX, MIN (범위 기반)</li>
            <li>ABS (단일 값)</li>
            <li>기타 Excel 내장 함수들</li>
            <li>수식을 셀에 직접 입력 (상대 참조 자동 처리)</li>
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
              placeholder="예: A1:A5"
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

        {/* 범위 기반 함수 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>📈 범위 기반 함수</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={calculateSum}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              SUM
            </button>
            <button
              onClick={calculateAverage}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              AVERAGE
            </button>
            <button
              onClick={calculateCount}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              COUNT
            </button>
            <button
              onClick={calculateMax}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              MAX
            </button>
            <button
              onClick={calculateMin}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              MIN
            </button>
          </div>
        </div>

        {/* 단일 값 함수 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔢 단일 값 함수</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>숫자:</label>
            <input
              type="number"
              value={functionArgs}
              onChange={(e) => setFunctionArgs(e.target.value)}
              placeholder="예: -5"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <button
            onClick={calculateAbs}
            style={{
              padding: "8px 16px",
              backgroundColor: "#9c27b0",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            ABS (절댓값)
          </button>
        </div>

        {/* 수식 입력 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #e91e63" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#e91e63" }}>📝 수식 입력 (셀에 넣기)</h4>
          <p style={{ margin: "0 0 15px 0", fontSize: "12px", color: "#666" }}>
            원하는 범위에 수식을 입력합니다. Excel이 자동으로 각 셀에 상대 참조를 적용합니다.
            <br />
            예: C1:C10에 =A1+B1을 넣으면, C1에는 =A1+B1, C2에는 =A2+B2가 자동으로 적용됩니다.
          </p>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>대상 범위 (수식을 넣을 범위):</label>
            <input
              type="text"
              value={formulaTargetRange}
              onChange={(e) => setFormulaTargetRange(e.target.value)}
              placeholder="예: C1:C10"
              disabled={useSelectionForFormula}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px", opacity: useSelectionForFormula ? 0.6 : 1 }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "flex", alignItems: "center", fontSize: "12px" }}>
              <input
                type="checkbox"
                checked={useSelectionForFormula}
                onChange={(e) => {
                  setUseSelectionForFormula(e.target.checked);
                  if (e.target.checked) {
                    getSelectedRangeForFormula();
                  }
                }}
                style={{ marginRight: "8px" }}
              />
              선택된 범위 사용
            </label>
          </div>
          {!useSelectionForFormula && (
            <button
              onClick={getSelectedRangeForFormula}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                marginBottom: "10px",
              }}
            >
              현재 선택된 범위 가져오기
            </button>
          )}
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>수식:</label>
            <input
              type="text"
              value={formulaExpression}
              onChange={(e) => setFormulaExpression(e.target.value)}
              placeholder="예: =A1+B1 또는 A1+B1 (자동으로 = 추가)"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px", fontFamily: "monospace" }}
            />
            <div style={{ fontSize: "11px", color: "#666", marginBottom: "10px" }}>
              예시: =A1+B1, =SUM(A1:A10), =IF(A1>0,"양수","음수"), =VLOOKUP(A1,Sheet2!A:B,2,FALSE)
            </div>
          </div>
          <button
            onClick={insertFormula}
            style={{
              padding: "8px 16px",
              backgroundColor: "#e91e63",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              width: "100%",
            }}
          >
            수식 입력
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
          {result || "위 버튼을 클릭하여 Excel 함수 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Functions;
