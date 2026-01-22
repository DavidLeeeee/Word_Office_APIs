import React, { useState } from "react";

/* global Excel */

const Validation: React.FC = () => {
  const [result, setResult] = useState("");
  const [rangeAddress, setRangeAddress] = useState("A1");
  const [useSelection, setUseSelection] = useState(false);
  const [validationType, setValidationType] = useState<"WholeNumber" | "Decimal" | "List" | "Date" | "Time" | "TextLength">("WholeNumber");
  const [operator, setOperator] = useState<"Between" | "NotBetween" | "EqualTo" | "NotEqualTo" | "GreaterThan" | "LessThan" | "GreaterThanOrEqualTo" | "LessThanOrEqualTo">("Between");
  const [formula1, setFormula1] = useState("0");
  const [formula2, setFormula2] = useState("100");
  const [listSource, setListSource] = useState("옵션1,옵션2,옵션3");
  const [showDropdown, setShowDropdown] = useState(true);
  const [ignoreBlanks, setIgnoreBlanks] = useState(true);
  const [errorTitle, setErrorTitle] = useState("오류");
  const [errorMessage, setErrorMessage] = useState("입력한 값이 유효하지 않습니다.");
  const [promptTitle, setPromptTitle] = useState("입력 안내");
  const [promptMessage, setPromptMessage] = useState("");

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

  // 1. 데이터 검증 정보 읽기
  const readValidation = async () => {
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
        const validation = range.dataValidation;
        validation.load("type,ignoreBlanks,valid");
        const errorAlert = validation.errorAlert;
        errorAlert.load("message,showAlert,style,title");
        const prompt = validation.prompt;
        prompt.load("message,showPrompt,title");
        await context.sync();

        let info = `데이터 검증 정보:\n\n범위: ${range.address}\n타입: ${validation.type}\n빈 셀 무시: ${validation.ignoreBlanks ? "예" : "아니오"}\n유효성: ${validation.valid === null ? "혼합" : (validation.valid ? "유효" : "무효")}\n\n오류 알림:\n제목: ${errorAlert.title || "(없음)"}\n메시지: ${errorAlert.message || "(없음)"}\n표시: ${errorAlert.showAlert ? "예" : "아니오"}\n스타일: ${errorAlert.style || "(없음)"}\n\n입력 프롬프트:\n제목: ${prompt.title || "(없음)"}\n메시지: ${prompt.message || "(없음)"}\n표시: ${prompt.showPrompt ? "예" : "아니오"}`;

        // 규칙 정보 읽기
        if (validation.type !== "None") {
          const rule = validation.rule;
          rule.load();
          await context.sync();

          if (validation.type === "List" && rule.list) {
            info += `\n\n규칙 (목록):\n소스: ${rule.list.source || "(없음)"}\n드롭다운 표시: ${rule.list.inCellDropDown ? "예" : "아니오"}`;
          } else if ((validation.type === "WholeNumber" || validation.type === "Decimal" || validation.type === "TextLength") && (rule.wholeNumber || rule.decimal || rule.textLength)) {
            const basicRule = rule.wholeNumber || rule.decimal || rule.textLength;
            if (basicRule) {
              info += `\n\n규칙:\n연산자: ${basicRule.operator}\n값1: ${basicRule.formula1}\n값2: ${basicRule.formula2 || "(없음)"}`;
            }
          } else if ((validation.type === "Date" || validation.type === "Time") && (rule.date || rule.time)) {
            const dateRule = rule.date || rule.time;
            if (dateRule) {
              info += `\n\n규칙:\n연산자: ${dateRule.operator}\n값1: ${dateRule.formula1}\n값2: ${dateRule.formula2 || "(없음)"}`;
            }
          }
        }

        info += `\n\n과정:\n1. range.dataValidation으로 검증 객체 가져오기\n2. validation.load()로 속성 로드\n3. validation.rule로 규칙 정보 로드\n4. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 2. 데이터 검증 설정 (정수)
  const setWholeNumberValidation = async () => {
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
        const validation = range.dataValidation;
        
        validation.rule = {
          wholeNumber: {
            formula1: formula1,
            formula2: operator === "Between" || operator === "NotBetween" ? formula2 : undefined,
            operator: operator
          }
        };
        validation.ignoreBlanks = ignoreBlanks;
        validation.errorAlert.title = errorTitle;
        validation.errorAlert.message = errorMessage;
        validation.errorAlert.showAlert = true;
        validation.errorAlert.style = "Stop";
        if (promptMessage.trim()) {
          validation.prompt.title = promptTitle;
          validation.prompt.message = promptMessage;
          validation.prompt.showPrompt = true;
        }
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`정수 데이터 검증 설정 완료!\n범위: ${range.address}\n타입: 정수\n연산자: ${operator}\n값1: ${formula1}\n값2: ${operator === "Between" || operator === "NotBetween" ? formula2 : "(사용 안 함)"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.rule.wholeNumber 설정\n3. validation.errorAlert, prompt 설정\n4. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 3. 데이터 검증 설정 (소수)
  const setDecimalValidation = async () => {
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
        const validation = range.dataValidation;
        
        validation.rule = {
          decimal: {
            formula1: formula1,
            formula2: operator === "Between" || operator === "NotBetween" ? formula2 : undefined,
            operator: operator
          }
        };
        validation.ignoreBlanks = ignoreBlanks;
        validation.errorAlert.title = errorTitle;
        validation.errorAlert.message = errorMessage;
        validation.errorAlert.showAlert = true;
        validation.errorAlert.style = "Stop";
        if (promptMessage.trim()) {
          validation.prompt.title = promptTitle;
          validation.prompt.message = promptMessage;
          validation.prompt.showPrompt = true;
        }
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`소수 데이터 검증 설정 완료!\n범위: ${range.address}\n타입: 소수\n연산자: ${operator}\n값1: ${formula1}\n값2: ${operator === "Between" || operator === "NotBetween" ? formula2 : "(사용 안 함)"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.rule.decimal 설정\n3. validation.errorAlert, prompt 설정\n4. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 4. 데이터 검증 설정 (목록)
  const setListValidation = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    if (!listSource.trim()) {
      setResult("목록 소스를 입력해주세요 (쉼표로 구분).");
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
        const validation = range.dataValidation;
        
        validation.rule = {
          list: {
            source: listSource,
            inCellDropDown: showDropdown
          }
        };
        validation.ignoreBlanks = ignoreBlanks;
        validation.errorAlert.title = errorTitle;
        validation.errorAlert.message = errorMessage;
        validation.errorAlert.showAlert = true;
        validation.errorAlert.style = "Stop";
        if (promptMessage.trim()) {
          validation.prompt.title = promptTitle;
          validation.prompt.message = promptMessage;
          validation.prompt.showPrompt = true;
        }
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`목록 데이터 검증 설정 완료!\n범위: ${range.address}\n타입: 목록\n소스: ${listSource}\n드롭다운 표시: ${showDropdown ? "예" : "아니오"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.rule.list 설정\n3. validation.errorAlert, prompt 설정\n4. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 5. 데이터 검증 설정 (날짜)
  const setDateValidation = async () => {
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
        const validation = range.dataValidation;
        
        const date1 = new Date(formula1);
        const date2 = operator === "Between" || operator === "NotBetween" ? new Date(formula2) : undefined;
        
        validation.rule = {
          date: {
            formula1: date1,
            formula2: date2,
            operator: operator
          }
        };
        validation.ignoreBlanks = ignoreBlanks;
        validation.errorAlert.title = errorTitle;
        validation.errorAlert.message = errorMessage;
        validation.errorAlert.showAlert = true;
        validation.errorAlert.style = "Stop";
        if (promptMessage.trim()) {
          validation.prompt.title = promptTitle;
          validation.prompt.message = promptMessage;
          validation.prompt.showPrompt = true;
        }
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`날짜 데이터 검증 설정 완료!\n범위: ${range.address}\n타입: 날짜\n연산자: ${operator}\n값1: ${formula1}\n값2: ${operator === "Between" || operator === "NotBetween" ? formula2 : "(사용 안 함)"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.rule.date 설정\n3. validation.errorAlert, prompt 설정\n4. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 데이터 검증 설정 (텍스트 길이)
  const setTextLengthValidation = async () => {
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
        const validation = range.dataValidation;
        
        validation.rule = {
          textLength: {
            formula1: formula1,
            formula2: operator === "Between" || operator === "NotBetween" ? formula2 : undefined,
            operator: operator
          }
        };
        validation.ignoreBlanks = ignoreBlanks;
        validation.errorAlert.title = errorTitle;
        validation.errorAlert.message = errorMessage;
        validation.errorAlert.showAlert = true;
        validation.errorAlert.style = "Stop";
        if (promptMessage.trim()) {
          validation.prompt.title = promptTitle;
          validation.prompt.message = promptMessage;
          validation.prompt.showPrompt = true;
        }
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`텍스트 길이 데이터 검증 설정 완료!\n범위: ${range.address}\n타입: 텍스트 길이\n연산자: ${operator}\n값1: ${formula1}\n값2: ${operator === "Between" || operator === "NotBetween" ? formula2 : "(사용 안 함)"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.rule.textLength 설정\n3. validation.errorAlert, prompt 설정\n4. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 7. 데이터 검증 제거
  const clearValidation = async () => {
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
        const validation = range.dataValidation;
        validation.clear();
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`데이터 검증 제거 완료!\n범위: ${range.address}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.clear()로 검증 제거\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 8. 무효한 셀 찾기
  const getInvalidCells = async () => {
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
        const validation = range.dataValidation;
        const invalidCells = validation.getInvalidCellsOrNullObject();
        invalidCells.load("address");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        if (invalidCells.isNullObject) {
          setResult(`무효한 셀 검사 완료!\n범위: ${range.address}\n결과: 모든 셀이 유효합니다.\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.getInvalidCellsOrNullObject()로 무효한 셀 찾기\n3. context.sync()로 동기화`);
        } else {
          setResult(`무효한 셀 검사 완료!\n범위: ${range.address}\n무효한 셀: ${invalidCells.address}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.dataValidation.getInvalidCellsOrNullObject()로 무효한 셀 찾기\n3. context.sync()로 동기화`);
        }
      });
    } catch (error: any) {
      if (error.code === "ItemNotFound") {
        setResult(`무효한 셀 검사 완료!\n결과: 모든 셀이 유효합니다.`);
      } else {
        setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
      }
    }
  };

  // 검증 타입에 따라 적절한 함수 호출
  const applyValidation = async () => {
    switch (validationType) {
      case "WholeNumber":
        await setWholeNumberValidation();
        break;
      case "Decimal":
        await setDecimalValidation();
        break;
      case "List":
        await setListValidation();
        break;
      case "Date":
        await setDateValidation();
        break;
      case "Time":
        await setDateValidation(); // Time도 DateTimeDataValidation 사용
        break;
      case "TextLength":
        await setTextLengthValidation();
        break;
      default:
        setResult("지원되지 않는 검증 타입입니다.");
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 데이터 검증</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>✅ Excel 데이터 검증 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 데이터 검증은 셀에 입력되는 값을 제한하고 검증하는 기능입니다. 잘못된 데이터 입력을 방지할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>정수 검증 (WholeNumber)</li>
            <li>소수 검증 (Decimal)</li>
            <li>목록 검증 (List)</li>
            <li>날짜 검증 (Date)</li>
            <li>시간 검증 (Time)</li>
            <li>텍스트 길이 검증 (TextLength)</li>
            <li>검증 정보 읽기</li>
            <li>검증 제거</li>
            <li>무효한 셀 찾기</li>
          </ul>
        </div>

        {/* 범위 지정 */}
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
        </div>

        {/* 검증 정보 읽기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📖 검증 정보</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={readValidation}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              검증 정보 읽기
            </button>
            <button
              onClick={getInvalidCells}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              무효한 셀 찾기
            </button>
            <button
              onClick={clearValidation}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              검증 제거
            </button>
          </div>
        </div>

        {/* 검증 설정 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 검증 설정</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>검증 타입:</label>
            <select
              value={validationType}
              onChange={(e) => setValidationType(e.target.value as any)}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            >
              <option value="WholeNumber">정수</option>
              <option value="Decimal">소수</option>
              <option value="List">목록</option>
              <option value="Date">날짜</option>
              <option value="Time">시간</option>
              <option value="TextLength">텍스트 길이</option>
            </select>
          </div>

          {validationType === "List" ? (
            <>
              <div style={{ marginBottom: "10px" }}>
                <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>목록 소스 (쉼표로 구분):</label>
                <input
                  type="text"
                  value={listSource}
                  onChange={(e) => setListSource(e.target.value)}
                  placeholder="예: 옵션1,옵션2,옵션3"
                  style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
                />
              </div>
              <div style={{ marginBottom: "10px" }}>
                <label style={{ display: "flex", alignItems: "center", gap: "10px", cursor: "pointer" }}>
                  <input
                    type="checkbox"
                    checked={showDropdown}
                    onChange={(e) => setShowDropdown(e.target.checked)}
                  />
                  <span>셀에 드롭다운 표시</span>
                </label>
              </div>
            </>
          ) : (
            <>
              <div style={{ marginBottom: "10px" }}>
                <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>연산자:</label>
                <select
                  value={operator}
                  onChange={(e) => setOperator(e.target.value as any)}
                  style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
                >
                  <option value="Between">사이</option>
                  <option value="NotBetween">사이 아님</option>
                  <option value="EqualTo">같음</option>
                  <option value="NotEqualTo">같지 않음</option>
                  <option value="GreaterThan">보다 큼</option>
                  <option value="LessThan">보다 작음</option>
                  <option value="GreaterThanOrEqualTo">보다 크거나 같음</option>
                  <option value="LessThanOrEqualTo">보다 작거나 같음</option>
                </select>
              </div>
              <div style={{ marginBottom: "10px" }}>
                <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>값1:</label>
                <input
                  type="text"
                  value={formula1}
                  onChange={(e) => setFormula1(e.target.value)}
                  placeholder={validationType === "Date" || validationType === "Time" ? "예: 2024-01-01" : "예: 0"}
                  style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
                />
              </div>
              {(operator === "Between" || operator === "NotBetween") && (
                <div style={{ marginBottom: "10px" }}>
                  <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>값2:</label>
                  <input
                    type="text"
                    value={formula2}
                    onChange={(e) => setFormula2(e.target.value)}
                    placeholder={validationType === "Date" || validationType === "Time" ? "예: 2024-12-31" : "예: 100"}
                    style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
                  />
                </div>
              )}
            </>
          )}

          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "flex", alignItems: "center", gap: "10px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={ignoreBlanks}
                onChange={(e) => setIgnoreBlanks(e.target.checked)}
              />
              <span>빈 셀 무시</span>
            </label>
          </div>

          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>오류 제목:</label>
            <input
              type="text"
              value={errorTitle}
              onChange={(e) => setErrorTitle(e.target.value)}
              placeholder="예: 오류"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>오류 메시지:</label>
            <input
              type="text"
              value={errorMessage}
              onChange={(e) => setErrorMessage(e.target.value)}
              placeholder="예: 입력한 값이 유효하지 않습니다."
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>프롬프트 제목 (선택사항):</label>
            <input
              type="text"
              value={promptTitle}
              onChange={(e) => setPromptTitle(e.target.value)}
              placeholder="예: 입력 안내"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>프롬프트 메시지 (선택사항):</label>
            <input
              type="text"
              value={promptMessage}
              onChange={(e) => setPromptMessage(e.target.value)}
              placeholder="예: 0과 100 사이의 값을 입력하세요."
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>

          <button
            onClick={applyValidation}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            검증 설정 적용
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
          {result || "위 버튼을 클릭하여 Excel 데이터 검증 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Validation;
