import React, { useState } from "react";

/* global Word */

const Format: React.FC = () => {
  const [result, setResult] = useState("");
  
  // 텍스트 서식 관련 상태
  const [fontName, setFontName] = useState("Arial");
  const [fontSize, setFontSize] = useState("12");
  const [fontColor, setFontColor] = useState("#000000");
  const [highlightColor, setHighlightColor] = useState("#FFFF00");
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const [underline, setUnderline] = useState(false);
  const [strikethrough, setStrikethrough] = useState(false);
  
  // 문단 서식 관련 상태
  const [alignment, setAlignment] = useState<"Left" | "Centered" | "Right" | "Justified">("Left");
  const [leftIndent, setLeftIndent] = useState("0");
  const [rightIndent, setRightIndent] = useState("0");
  const [firstLineIndent, setFirstLineIndent] = useState("0");
  const [lineSpacing, setLineSpacing] = useState("1.0");
  const [beforeSpacing, setBeforeSpacing] = useState("0");
  const [afterSpacing, setAfterSpacing] = useState("0");

  // 1. 선택된 텍스트의 현재 서식 읽기
  const readCurrentFormat = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text,paragraphs");
        
        const font = selection.font;
        // 올바른 속성명: strikeThrough
        font.load("name,size,color,bold,italic,underline,strikeThrough,highlightColor");
        await context.sync();

        if (selection.text.trim() === "") {
          setResult("텍스트를 선택한 후 서식을 확인해주세요.");
          return;
        }

        // 문단 서식은 첫 문단 기준으로만 표시
        const paragraphs = selection.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraphs.items.length === 0) {
          setResult("문단을 찾을 수 없습니다.");
          return;
        }

        const firstPara = paragraphs.items[0];
        // Paragraph 객체에 직접 문단 서식 속성이 있음 (paragraphFormat이 아님)
        firstPara.load("alignment,leftIndent,rightIndent,firstLineIndent,lineSpacing,spaceBefore,spaceAfter");
        await context.sync();

        const formatInfo = `선택된 텍스트: "${selection.text.substring(0, 50)}${selection.text.length > 50 ? "..." : ""}"\n\n📝 텍스트 서식:\n글꼴: ${font.name}\n크기: ${font.size}pt\n색상: ${font.color}\n강조색: ${font.highlightColor || "없음"}\n굵게: ${font.bold ? "예" : "아니오"}\n이탤릭: ${font.italic ? "예" : "아니오"}\n밑줄: ${font.underline !== "None" ? font.underline : "없음"}\n취소선: ${font.strikeThrough ? "예" : "아니오"}\n\n📄 문단 서식:\n정렬: ${firstPara.alignment}\n왼쪽 들여쓰기: ${firstPara.leftIndent}pt\n오른쪽 들여쓰기: ${firstPara.rightIndent}pt\n첫 줄 들여쓰기: ${firstPara.firstLineIndent}pt\n줄 간격: ${firstPara.lineSpacing}\n문단 앞 간격: ${firstPara.spaceBefore}pt\n문단 뒤 간격: ${firstPara.spaceAfter}pt\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.font.load()로 글꼴 속성 로드\n3. selection.paragraphs.items[0].load()로 문단 서식 속성 로드\n4. context.sync()로 동기화`;

        setResult(formatInfo);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 텍스트 서식 적용
  const applyTextFormat = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (selection.text.trim() === "") {
          setResult("텍스트를 선택한 후 서식을 적용해주세요.");
          return;
        }

        const font = selection.font;
        font.name = fontName;
        font.size = parseFloat(fontSize);
        font.color = fontColor;
        font.bold = bold;
        font.italic = italic;
        font.underline = underline ? "Single" : "None";
        // 올바른 속성명: strikeThrough
        (font as any).strikeThrough = strikethrough;
        
        await context.sync();

        setResult(`텍스트 서식 적용 완료!\n\n적용된 서식:\n글꼴: ${fontName}\n크기: ${fontSize}pt\n색상: ${fontColor}\n굵게: ${bold ? "예" : "아니오"}\n이탤릭: ${italic ? "예" : "아니오"}\n밑줄: ${underline ? "예" : "아니오"}\n취소선: ${strikethrough ? "예" : "아니오"}\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.font로 글꼴 객체 접근\n3. font.name, size, color, bold, italic 등 설정\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 강조색 적용
  const applyHighlight = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (selection.text.trim() === "") {
          setResult("텍스트를 선택한 후 강조색을 적용해주세요.");
          return;
        }

        const font = selection.font;
        font.highlightColor = highlightColor;
        await context.sync();

        setResult(`강조색 적용 완료!\n색상: ${highlightColor}\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.font.highlightColor 설정\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 문단 서식 적용
  const applyParagraphFormat = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text,paragraphs");
        await context.sync();

        if (selection.text.trim() === "") {
          setResult("텍스트를 선택한 후 문단 서식을 적용해주세요.");
          return;
        }

        // 첫 번째 문단 가져오기
        const paragraphs = selection.paragraphs;
        paragraphs.load("items");
        await context.sync();

        if (paragraphs.items.length === 0) {
          setResult("문단을 찾을 수 없습니다.");
          return;
        }

        const firstPara = paragraphs.items[0];
        // Paragraph 객체에 직접 문단 서식 속성이 있음 (paragraphFormat이 아님)
        firstPara.load("alignment,leftIndent,rightIndent,firstLineIndent,lineSpacing,spaceBefore,spaceAfter");
        await context.sync();

        // 이제 속성 설정 가능
        firstPara.alignment = alignment;
        firstPara.leftIndent = parseFloat(leftIndent);
        firstPara.rightIndent = parseFloat(rightIndent);
        firstPara.firstLineIndent = parseFloat(firstLineIndent);
        firstPara.lineSpacing = parseFloat(lineSpacing);
        firstPara.spaceBefore = parseFloat(beforeSpacing);
        firstPara.spaceAfter = parseFloat(afterSpacing);

        await context.sync();

        setResult(
          `문단 서식 적용 완료!\n\n적용된 서식:\n정렬: ${alignment}\n왼쪽 들여쓰기: ${leftIndent}pt\n오른쪽 들여쓰기: ${rightIndent}pt\n첫 줄 들여쓰기: ${firstLineIndent}pt\n줄 간격: ${lineSpacing}\n문단 앞 간격: ${beforeSpacing}pt\n문단 뒤 간격: ${afterSpacing}pt\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.paragraphFormat으로 문단 서식 객체 가져오기\n3. paragraphFormat의 정렬/들여쓰기/간격 속성 설정\n4. context.sync()로 동기화`
        );
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 문서 여백 확인 (읽기 전용)
  const checkPageMargins = async () => {
    try {
      await Word.run(async (context) => {
        const sections = context.document.sections;
        sections.load("items/pageSetup/topMargin,items/pageSetup/bottomMargin,items/pageSetup/leftMargin,items/pageSetup/rightMargin,items/pageSetup/pageWidth,items/pageSetup/pageHeight");
        await context.sync();

        if (sections.items.length === 0) {
          setResult("문서에 섹션이 없습니다.");
          return;
        }

        const firstSection = sections.items[0];
        const pageSetup = firstSection.pageSetup;
        
        const marginInfo = `📄 문서 레이아웃 정보:\n\n여백:\n위쪽: ${pageSetup.topMargin}pt\n아래쪽: ${pageSetup.bottomMargin}pt\n왼쪽: ${pageSetup.leftMargin}pt\n오른쪽: ${pageSetup.rightMargin}pt\n\n페이지 크기:\n너비: ${pageSetup.pageWidth}pt\n높이: ${pageSetup.pageHeight}pt\n\n⚠️ 참고: Word JavaScript API에서는 페이지 여백과 크기를 읽을 수만 있고, 설정은 Word UI에서 해야 합니다.\n\n과정:\n1. context.document.sections로 섹션 컬렉션 가져오기\n2. section.pageSetup으로 페이지 설정 접근\n3. pageSetup.topMargin, bottomMargin 등 로드\n4. context.sync()로 동기화`;

        setResult(marginInfo);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: Word JavaScript API에서는 페이지 여백과 크기 설정이 제한적입니다.`);
    }
  };

  // 6. 빠른 서식 적용 (볼드/이탤릭/밑줄 토글)
  const toggleBold = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const font = selection.font;
        font.load("bold");
        await context.sync();
        font.bold = !font.bold;
        await context.sync();
        setResult(`굵게 ${font.bold ? "적용" : "해제"} 완료!`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  const toggleItalic = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const font = selection.font;
        font.load("italic");
        await context.sync();
        font.italic = !font.italic;
        await context.sync();
        setResult(`이탤릭 ${font.italic ? "적용" : "해제"} 완료!`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  const toggleUnderline = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const font = selection.font;
        font.load("underline");
        await context.sync();
        font.underline = font.underline === "None" ? "Single" : "None";
        await context.sync();
        setResult(`밑줄 ${font.underline !== "None" ? "적용" : "해제"} 완료!`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word 문서 서식</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#856404", fontSize: "14px" }}>📝 서식 기능 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#856404" }}>
            이 섹션에서는 Word 문서의 <strong>텍스트 서식</strong>과 <strong>문단 서식</strong>을 수정할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#856404", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#856404" }}>
            <li>텍스트 서식: 글꼴, 크기, 색상, 굵기, 이탤릭, 밑줄, 취소선, 강조색</li>
            <li>문단 서식: 정렬, 들여쓰기, 줄 간격, 문단 간격</li>
            <li>페이지 레이아웃 정보 확인 (읽기 전용)</li>
          </ul>
          <p style={{ margin: "8px 0", color: "#d32f2f", fontSize: "12px", fontStyle: "italic" }}>
            ⚠️ 제약사항: 자간(character spacing)과 페이지 여백/크기 설정은 Word JavaScript API에서 지원되지 않습니다.
          </p>
        </div>

        {/* 현재 서식 확인 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #2196f3" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>🔍 현재 서식 확인</h4>
          <button
            onClick={readCurrentFormat}
            style={{
              padding: "8px 16px",
              backgroundColor: "#2196f3",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            선택 영역 서식 확인
          </button>
          <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
            사용법: Word 문서에서 텍스트를 선택한 후 버튼을 클릭하세요.
          </div>
        </div>

        {/* 빠른 서식 토글 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>⚡ 빠른 서식</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={toggleBold}
              style={{
                padding: "8px 16px",
                backgroundColor: bold ? "#4caf50" : "#e0e0e0",
                color: bold ? "#fff" : "#000",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontWeight: "bold",
              }}
            >
              굵게
            </button>
            <button
              onClick={toggleItalic}
              style={{
                padding: "8px 16px",
                backgroundColor: italic ? "#4caf50" : "#e0e0e0",
                color: italic ? "#fff" : "#000",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontStyle: "italic",
              }}
            >
              이탤릭
            </button>
            <button
              onClick={toggleUnderline}
              style={{
                padding: "8px 16px",
                backgroundColor: underline ? "#4caf50" : "#e0e0e0",
                color: underline ? "#fff" : "#000",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                textDecoration: "underline",
              }}
            >
              밑줄
            </button>
          </div>
        </div>

        {/* 텍스트 서식 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>📝 텍스트 서식</h4>
          
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>글꼴</label>
              <input
                type="text"
                value={fontName}
                onChange={(e) => setFontName(e.target.value)}
                placeholder="예: Arial, 맑은 고딕"
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>크기 (pt)</label>
              <input
                type="number"
                value={fontSize}
                onChange={(e) => setFontSize(e.target.value)}
                min="1"
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>텍스트 색상</label>
              <input
                type="color"
                value={fontColor}
                onChange={(e) => setFontColor(e.target.value)}
                style={{ width: "100%", padding: "4px", border: "1px solid #ddd", borderRadius: "5px", height: "40px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>강조색</label>
              <input
                type="color"
                value={highlightColor}
                onChange={(e) => setHighlightColor(e.target.value)}
                style={{ width: "100%", padding: "4px", border: "1px solid #ddd", borderRadius: "5px", height: "40px" }}
              />
            </div>
          </div>

          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", flexWrap: "wrap" }}>
            <label style={{ display: "flex", alignItems: "center", gap: "5px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={bold}
                onChange={(e) => setBold(e.target.checked)}
              />
              <span>굵게</span>
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: "5px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={italic}
                onChange={(e) => setItalic(e.target.checked)}
              />
              <span>이탤릭</span>
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: "5px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={underline}
                onChange={(e) => setUnderline(e.target.checked)}
              />
              <span>밑줄</span>
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: "5px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={strikethrough}
                onChange={(e) => setStrikethrough(e.target.checked)}
              />
              <span>취소선</span>
            </label>
          </div>

          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={applyTextFormat}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              텍스트 서식 적용
            </button>
            <button
              onClick={applyHighlight}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              강조색만 적용
            </button>
          </div>
        </div>

        {/* 문단 서식 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff5722" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff5722" }}>📄 문단 서식</h4>
          
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>정렬</label>
            <select
              value={alignment}
              onChange={(e) => setAlignment(e.target.value as "Left" | "Centered" | "Right" | "Justified")}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            >
              <option value="Left">왼쪽</option>
              <option value="Centered">가운데</option>
              <option value="Right">오른쪽</option>
              <option value="Justified">양쪽</option>
            </select>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>왼쪽 들여쓰기 (pt)</label>
              <input
                type="number"
                value={leftIndent}
                onChange={(e) => setLeftIndent(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>오른쪽 들여쓰기 (pt)</label>
              <input
                type="number"
                value={rightIndent}
                onChange={(e) => setRightIndent(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>첫 줄 들여쓰기 (pt)</label>
              <input
                type="number"
                value={firstLineIndent}
                onChange={(e) => setFirstLineIndent(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>줄 간격</label>
              <input
                type="number"
                value={lineSpacing}
                onChange={(e) => setLineSpacing(e.target.value)}
                step="0.1"
                min="0.5"
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>문단 앞 간격 (pt)</label>
              <input
                type="number"
                value={beforeSpacing}
                onChange={(e) => setBeforeSpacing(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", marginBottom: "5px", fontSize: "13px" }}>문단 뒤 간격 (pt)</label>
              <input
                type="number"
                value={afterSpacing}
                onChange={(e) => setAfterSpacing(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
          </div>

          <button
            onClick={applyParagraphFormat}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff5722",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            문단 서식 적용
          </button>
        </div>

        {/* 페이지 레이아웃 확인 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>📐 페이지 레이아웃</h4>
          <button
            onClick={checkPageMargins}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            여백 및 페이지 크기 확인
          </button>
          <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
            ⚠️ 참고: 페이지 여백과 크기는 읽기만 가능하며, 설정은 Word UI에서 해야 합니다.
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
          {result || "위 버튼을 클릭하여 서식 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Format;
