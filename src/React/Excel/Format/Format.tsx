import React, { useState } from "react";

/* global Excel */

const Format: React.FC = () => {
  const [result, setResult] = useState("");
  const [rangeAddress, setRangeAddress] = useState("A1");
  const [useSelection, setUseSelection] = useState(false);

  // 폰트 서식
  const [fontName, setFontName] = useState("Calibri");
  const [fontSize, setFontSize] = useState("11");
  const [fontColor, setFontColor] = useState("#000000");
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [underline, setUnderline] = useState(false);
  const [strikethrough, setStrikethrough] = useState(false);

  // 채우기 서식
  const [fillColor, setFillColor] = useState("#FFFFFF");

  // 정렬
  const [horizontalAlignment, setHorizontalAlignment] = useState<"General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed">("General");
  const [verticalAlignment, setVerticalAlignment] = useState<"Top" | "Center" | "Bottom" | "Justify" | "Distributed">("Bottom");
  const [wrapText, setWrapText] = useState(false);

  // 숫자 서식
  const [numberFormat, setNumberFormat] = useState("General");

  // 행/열 크기
  const [columnWidth, setColumnWidth] = useState("");
  const [rowHeight, setRowHeight] = useState("");

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

  // 1. 현재 서식 읽기
  const readCurrentFormat = async () => {
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
        
        range.load("address,numberFormat");
        const format = range.format;
        format.load("horizontalAlignment,verticalAlignment,wrapText,columnWidth,rowHeight");
        const font = format.font;
        font.load("name,size,color,bold,italic,underline,strikethrough");
        const fill = format.fill;
        fill.load("color");
        await context.sync();

        const formatInfo = `범위 서식 읽기 완료!\n주소: ${range.address}\n\n📝 폰트 서식:\n글꼴: ${font.name}\n크기: ${font.size}pt\n색상: ${font.color}\n굵게: ${font.bold ? "예" : "아니오"}\n이탤릭: ${font.italic ? "예" : "아니오"}\n밑줄: ${font.underline !== "None" ? font.underline : "없음"}\n취소선: ${font.strikethrough ? "예" : "아니오"}\n\n🎨 채우기 서식:\n배경색: ${fill.color || "없음"}\n\n📐 정렬:\n가로 정렬: ${format.horizontalAlignment}\n세로 정렬: ${format.verticalAlignment}\n자동 줄바꿈: ${format.wrapText ? "예" : "아니오"}\n\n🔢 숫자 서식:\n형식: ${range.numberFormat}\n\n📏 크기:\n열 너비: ${format.columnWidth || "표준"}\n행 높이: ${format.rowHeight || "표준"}\n\n과정:\n1. range.format으로 서식 객체 가져오기\n2. format.font, format.fill, format 속성 로드\n3. context.sync()로 동기화`;

        setResult(formatInfo);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 폰트 서식 적용
  const applyFontFormat = async () => {
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
        const font = range.format.font;
        font.name = fontName;
        font.size = parseFloat(fontSize);
        font.color = fontColor;
        font.bold = bold;
        font.italic = italic;
        font.underline = underline ? "Single" : "None";
        font.strikethrough = strikethrough;
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`폰트 서식 적용 완료!\n주소: ${range.address}\n\n적용된 서식:\n글꼴: ${fontName}\n크기: ${fontSize}pt\n색상: ${fontColor}\n굵게: ${bold ? "예" : "아니오"}\n이탤릭: ${italic ? "예" : "아니오"}\n밑줄: ${underline ? "예" : "아니오"}\n취소선: ${strikethrough ? "예" : "아니오"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.format.font로 폰트 객체 접근\n3. font.name, size, color, bold, italic 등 설정\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 채우기 서식 적용
  const applyFillFormat = async () => {
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
        const fill = range.format.fill;
        fill.color = fillColor;
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`채우기 서식 적용 완료!\n주소: ${range.address}\n배경색: ${fillColor}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.format.fill로 채우기 객체 접근\n3. fill.color 설정\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 정렬 적용
  const applyAlignment = async () => {
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
        const format = range.format;
        format.horizontalAlignment = horizontalAlignment;
        format.verticalAlignment = verticalAlignment;
        format.wrapText = wrapText;
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`정렬 적용 완료!\n주소: ${range.address}\n\n적용된 정렬:\n가로 정렬: ${horizontalAlignment}\n세로 정렬: ${verticalAlignment}\n자동 줄바꿈: ${wrapText ? "예" : "아니오"}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.format로 서식 객체 접근\n3. format.horizontalAlignment, verticalAlignment, wrapText 설정\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 숫자 서식 적용
  const applyNumberFormat = async () => {
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
        range.numberFormat = [[numberFormat]];
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`숫자 서식 적용 완료!\n주소: ${range.address}\n숫자 형식: ${numberFormat}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.numberFormat 설정\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. 열 너비 설정
  const applyColumnWidth = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }
    if (!columnWidth.trim()) {
      setResult("열 너비를 입력해주세요.");
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
        range.format.columnWidth = parseFloat(columnWidth);
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`열 너비 설정 완료!\n주소: ${range.address}\n열 너비: ${columnWidth}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.format.columnWidth 설정\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 7. 행 높이 설정
  const applyRowHeight = async () => {
    if (!useSelection && !rangeAddress.trim()) {
      setResult("범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }
    if (!rowHeight.trim()) {
      setResult("행 높이를 입력해주세요.");
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
        range.format.rowHeight = parseFloat(rowHeight);
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`행 높이 설정 완료!\n주소: ${range.address}\n행 높이: ${rowHeight}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.format.rowHeight 설정\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 열 너비 자동 맞춤
  const autofitColumns = async () => {
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
        range.format.autofitColumns();
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${rangeAddress}")`;
        setResult(`열 너비 자동 맞춤 완료!\n주소: ${range.address}\n\n과정:\n1. ${method}로 범위 가져오기\n2. range.format.autofitColumns() 호출\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 서식</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📝 Excel 서식 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 셀의 서식을 설정하고 관리할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>폰트 서식 (글꼴, 크기, 색상, 굵게, 이탤릭, 밑줄, 취소선)</li>
            <li>채우기 서식 (배경색)</li>
            <li>정렬 (가로/세로 정렬, 자동 줄바꿈)</li>
            <li>숫자 서식</li>
            <li>행/열 크기 조정</li>
            <li>열 너비 자동 맞춤</li>
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

        {/* 서식 읽기 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📖 서식 읽기</h4>
          <button
            onClick={readCurrentFormat}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            현재 서식 읽기
          </button>
        </div>

        {/* 폰트 서식 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>🔤 폰트 서식</h4>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>글꼴:</label>
              <input
                type="text"
                value={fontName}
                onChange={(e) => setFontName(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>크기:</label>
              <input
                type="number"
                value={fontSize}
                onChange={(e) => setFontSize(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>색상:</label>
              <input
                type="color"
                value={fontColor}
                onChange={(e) => setFontColor(e.target.value)}
                style={{ width: "100%", padding: "4px", border: "1px solid #ddd", borderRadius: "5px", height: "40px" }}
              />
            </div>
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={() => setBold(!bold)}
              style={{
                padding: "8px 16px",
                backgroundColor: bold ? "#ff9800" : "#ddd",
                color: bold ? "#fff" : "#000",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontWeight: bold ? "bold" : "normal",
              }}
            >
              굵게
            </button>
            <button
              onClick={() => setItalic(!italic)}
              style={{
                padding: "8px 16px",
                backgroundColor: italic ? "#ff9800" : "#ddd",
                color: italic ? "#fff" : "#000",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontStyle: italic ? "italic" : "normal",
              }}
            >
              이탤릭
            </button>
            <button
              onClick={() => setUnderline(!underline)}
              style={{
                padding: "8px 16px",
                backgroundColor: underline ? "#ff9800" : "#ddd",
                color: underline ? "#fff" : "#000",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                textDecoration: underline ? "underline" : "none",
              }}
            >
              밑줄
            </button>
            <button
              onClick={() => setStrikethrough(!strikethrough)}
              style={{
                padding: "8px 16px",
                backgroundColor: strikethrough ? "#ff9800" : "#ddd",
                color: strikethrough ? "#fff" : "#000",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                textDecoration: strikethrough ? "line-through" : "none",
              }}
            >
              취소선
            </button>
          </div>
          <button
            onClick={applyFontFormat}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            폰트 서식 적용
          </button>
        </div>

        {/* 채우기 서식 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🎨 채우기 서식</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", marginBottom: "10px" }}>
            <label style={{ fontSize: "13px" }}>배경색:</label>
            <input
              type="color"
              value={fillColor}
              onChange={(e) => setFillColor(e.target.value)}
              style={{ padding: "4px", border: "1px solid #ddd", borderRadius: "5px", height: "40px" }}
            />
          </div>
          <button
            onClick={applyFillFormat}
            style={{
              padding: "8px 16px",
              backgroundColor: "#9c27b0",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            채우기 서식 적용
          </button>
        </div>

        {/* 정렬 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>📐 정렬</h4>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>가로 정렬:</label>
              <select
                value={horizontalAlignment}
                onChange={(e) => setHorizontalAlignment(e.target.value as any)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              >
                <option value="General">일반</option>
                <option value="Left">왼쪽</option>
                <option value="Center">가운데</option>
                <option value="Right">오른쪽</option>
                <option value="Fill">채우기</option>
                <option value="Justify">양쪽 맞춤</option>
                <option value="CenterAcrossSelection">선택 영역 가운데</option>
                <option value="Distributed">분산</option>
              </select>
            </div>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>세로 정렬:</label>
              <select
                value={verticalAlignment}
                onChange={(e) => setVerticalAlignment(e.target.value as any)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              >
                <option value="Top">위</option>
                <option value="Center">가운데</option>
                <option value="Bottom">아래</option>
                <option value="Justify">양쪽 맞춤</option>
                <option value="Distributed">분산</option>
              </select>
            </div>
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "flex", alignItems: "center", gap: "10px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={wrapText}
                onChange={(e) => setWrapText(e.target.checked)}
              />
              <span>자동 줄바꿈</span>
            </label>
          </div>
          <button
            onClick={applyAlignment}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            정렬 적용
          </button>
        </div>

        {/* 숫자 서식 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #f44336" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#f44336" }}>🔢 숫자 서식</h4>
          <div style={{ marginBottom: "10px" }}>
            <select
              value={numberFormat}
              onChange={(e) => setNumberFormat(e.target.value)}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            >
              <option value="General">일반</option>
              <option value="0">숫자 (0)</option>
              <option value="0.00">숫자 (0.00)</option>
              <option value="#,##0">천 단위 구분 기호</option>
              <option value="0%">백분율 (0%)</option>
              <option value="0.00%">백분율 (0.00%)</option>
              <option value="mm/dd/yyyy">날짜 (mm/dd/yyyy)</option>
              <option value="hh:mm:ss">시간 (hh:mm:ss)</option>
              <option value="Currency">통화</option>
              <option value="Accounting">회계</option>
            </select>
            <input
              type="text"
              value={numberFormat}
              onChange={(e) => setNumberFormat(e.target.value)}
              placeholder="또는 사용자 지정 형식 입력 (예: 0.00%)"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
          </div>
          <button
            onClick={applyNumberFormat}
            style={{
              padding: "8px 16px",
              backgroundColor: "#f44336",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            숫자 서식 적용
          </button>
        </div>

        {/* 행/열 크기 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #e91e63" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#e91e63" }}>📏 행/열 크기</h4>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>열 너비:</label>
              <div style={{ display: "flex", gap: "5px" }}>
                <input
                  type="number"
                  value={columnWidth}
                  onChange={(e) => setColumnWidth(e.target.value)}
                  placeholder="예: 10"
                  style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
                />
                <button
                  onClick={applyColumnWidth}
                  style={{
                    padding: "8px 16px",
                    backgroundColor: "#e91e63",
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
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>행 높이:</label>
              <div style={{ display: "flex", gap: "5px" }}>
                <input
                  type="number"
                  value={rowHeight}
                  onChange={(e) => setRowHeight(e.target.value)}
                  placeholder="예: 20"
                  style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
                />
                <button
                  onClick={applyRowHeight}
                  style={{
                    padding: "8px 16px",
                    backgroundColor: "#e91e63",
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
          </div>
          <button
            onClick={autofitColumns}
            style={{
              padding: "8px 16px",
              backgroundColor: "#e91e63",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            열 너비 자동 맞춤
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
          {result || "위 버튼을 클릭하여 Excel 서식 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Format;
