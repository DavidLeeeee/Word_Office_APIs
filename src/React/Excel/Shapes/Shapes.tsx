import React, { useState } from "react";

/* global Excel */

const Shapes: React.FC = () => {
  const [result, setResult] = useState("");
  const [shapeName, setShapeName] = useState("");
  const [shapeType, setShapeType] = useState<"Rectangle" | "Ellipse" | "Triangle" | "Line" | "Image" | "TextBox">("Rectangle");
  const [imageBase64, setImageBase64] = useState("");
  const [textBoxText, setTextBoxText] = useState("텍스트");
  const [lineStartLeft, setLineStartLeft] = useState("100");
  const [lineStartTop, setLineStartTop] = useState("100");
  const [lineEndLeft, setLineEndLeft] = useState("200");
  const [lineEndTop, setLineEndTop] = useState("200");

  // 1. 도형 목록 가져오기
  const listShapes = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        shapes.load("items/name,items/id,items/geometricShapeType,items/width,items/height");
        await context.sync();

        if (shapes.items.length === 0) {
          setResult("현재 워크시트에 도형이 없습니다.");
          return;
        }

        let resultText = `도형 목록 (${shapes.items.length}개):\n\n`;
        shapes.items.forEach((shape, index) => {
          resultText += `${index + 1}. ${shape.name}\n`;
          resultText += `   ID: ${shape.id}\n`;
          resultText += `   타입: ${shape.geometricShapeType || "기타"}\n`;
          resultText += `   크기: ${shape.width}pt × ${shape.height}pt\n\n`;
        });

        resultText += `과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.shapes로 도형 컬렉션 가져오기\n3. shapes.load("items/name,items/id,...")로 속성 로드\n4. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 기하학적 도형 생성
  const createGeometricShape = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        const newShape = shapes.addGeometricShape(shapeType);
        newShape.load("name,id,geometricShapeType,width,height");
        await context.sync();

        setResult(`도형 생성 완료!\n도형 이름: ${newShape.name}\nID: ${newShape.id}\n타입: ${newShape.geometricShapeType}\n크기: ${newShape.width}pt × ${newShape.height}pt\n\n과정:\n1. sheet.shapes.addGeometricShape("${shapeType}")로 도형 생성\n2. newShape.load()로 속성 로드\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 3. 이미지 추가
  const addImage = async () => {
    if (!imageBase64.trim()) {
      setResult("Base64 이미지 데이터를 입력해주세요.\n\n참고: 이미지를 Base64로 변환해야 합니다. (data:image/png;base64,... 형식도 가능)");
      return;
    }

    try {
      // data:image/png;base64, 형식 제거
      let base64Data = imageBase64.trim();
      if (base64Data.includes(",")) {
        base64Data = base64Data.split(",")[1];
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        const newShape = shapes.addImage(base64Data);
        newShape.load("name,id,width,height");
        await context.sync();

        setResult(`이미지 추가 완료!\n도형 이름: ${newShape.name}\nID: ${newShape.id}\n크기: ${newShape.width}pt × ${newShape.height}pt\n\n과정:\n1. sheet.shapes.addImage(base64String)로 이미지 추가\n2. newShape.load()로 속성 로드\n3. context.sync()로 동기화\n\n참고: 이미지는 Base64 인코딩된 JPEG 또는 PNG 형식이어야 합니다.`);
        setImageBase64("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: Base64 형식이 올바른지 확인해주세요.`);
    }
  };

  // 4. 선 추가
  const addLine = async () => {
    const startLeft = parseFloat(lineStartLeft) || 100;
    const startTop = parseFloat(lineStartTop) || 100;
    const endLeft = parseFloat(lineEndLeft) || 200;
    const endTop = parseFloat(lineEndTop) || 200;

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        const newShape = shapes.addLine(startLeft, startTop, endLeft, endTop);
        newShape.load("name,id,width,height");
        await context.sync();

        setResult(`선 추가 완료!\n도형 이름: ${newShape.name}\nID: ${newShape.id}\n크기: ${newShape.width}pt × ${newShape.height}pt\n시작: (${startLeft}pt, ${startTop}pt)\n끝: (${endLeft}pt, ${endTop}pt)\n\n과정:\n1. sheet.shapes.addLine(${startLeft}, ${startTop}, ${endLeft}, ${endTop})로 선 추가\n2. newShape.load()로 속성 로드\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 5. 텍스트박스 추가
  const addTextBox = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        const newShape = shapes.addTextBox(textBoxText);
        newShape.load("name,id,width,height");
        await context.sync();

        setResult(`텍스트박스 추가 완료!\n도형 이름: ${newShape.name}\nID: ${newShape.id}\n크기: ${newShape.width}pt × ${newShape.height}pt\n텍스트: ${textBoxText}\n\n과정:\n1. sheet.shapes.addTextBox("${textBoxText}")로 텍스트박스 추가\n2. newShape.load()로 속성 로드\n3. context.sync()로 동기화`);
        setTextBoxText("텍스트");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 도형 정보 읽기
  const getShapeInfo = async () => {
    if (!shapeName.trim()) {
      setResult("도형 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        const shape = shapes.getItem(shapeName);
        
        shape.load("name,id,geometricShapeType,width,height,left,top,altTextTitle,altTextDescription");
        await context.sync();

        const info = `도형 정보:\n\n이름: ${shape.name}\nID: ${shape.id}\n타입: ${shape.geometricShapeType || "기타"}\n위치: (${shape.left}pt, ${shape.top}pt)\n크기: ${shape.width}pt × ${shape.height}pt\n대체 텍스트 제목: ${shape.altTextTitle || "(없음)"}\n대체 텍스트 설명: ${shape.altTextDescription || "(없음)"}\n\n과정:\n1. sheet.shapes.getItem("${shapeName}")로 도형 가져오기\n2. shape.load()로 속성 로드\n3. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 도형 이름을 확인해주세요.`);
    }
  };

  // 7. 도형 삭제
  const deleteShape = async () => {
    if (!shapeName.trim()) {
      setResult("삭제할 도형 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = sheet.shapes;
        const shape = shapes.getItem(shapeName);
        shape.load("name");
        await context.sync();

        const deletedName = shape.name;
        shape.delete();
        await context.sync();

        setResult(`도형 삭제 완료!\n삭제된 도형: ${deletedName}\n\n과정:\n1. sheet.shapes.getItem("${shapeName}")로 도형 가져오기\n2. shape.delete()로 도형 삭제\n3. context.sync()로 동기화`);
        setShapeName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 도형</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>🎨 Excel 도형 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 도형은 워크시트에 시각적 요소를 추가하는 기능입니다. 기하학적 도형, 이미지, 선, 텍스트박스를 추가할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>기하학적 도형 생성 (사각형, 원, 삼각형 등)</li>
            <li>이미지 추가 (Base64)</li>
            <li>선 추가</li>
            <li>텍스트박스 추가</li>
            <li>도형 목록 조회</li>
            <li>도형 정보 읽기</li>
            <li>도형 삭제</li>
          </ul>
        </div>

        {/* 도형 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 도형 목록</h4>
          <button
            onClick={listShapes}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            도형 목록 보기
          </button>
        </div>

        {/* 기하학적 도형 생성 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 기하학적 도형 생성</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>도형 타입:</label>
            <select
              value={shapeType}
              onChange={(e) => setShapeType(e.target.value as any)}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            >
              <option value="Rectangle">사각형</option>
              <option value="Ellipse">타원</option>
              <option value="Triangle">삼각형</option>
              <option value="Diamond">다이아몬드</option>
              <option value="RoundRectangle">둥근 사각형</option>
              <option value="Star5">별 (5각)</option>
              <option value="Heart">하트</option>
            </select>
          </div>
          <button
            onClick={createGeometricShape}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            도형 생성
          </button>
        </div>

        {/* 이미지 추가 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🖼️ 이미지 추가</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>Base64 이미지 데이터:</label>
            <textarea
              value={imageBase64}
              onChange={(e) => setImageBase64(e.target.value)}
              placeholder="Base64 인코딩된 이미지 데이터 입력 (data:image/png;base64,... 형식도 가능)"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minHeight: "80px", fontFamily: "monospace", fontSize: "11px" }}
            />
          </div>
          <button
            onClick={addImage}
            style={{
              padding: "8px 16px",
              backgroundColor: "#9c27b0",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            이미지 추가
          </button>
        </div>

        {/* 선 추가 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>📏 선 추가</h4>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginBottom: "10px" }}>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>시작 X (pt):</label>
              <input
                type="number"
                value={lineStartLeft}
                onChange={(e) => setLineStartLeft(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>시작 Y (pt):</label>
              <input
                type="number"
                value={lineStartTop}
                onChange={(e) => setLineStartTop(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>끝 X (pt):</label>
              <input
                type="number"
                value={lineEndLeft}
                onChange={(e) => setLineEndLeft(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
            <div>
              <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>끝 Y (pt):</label>
              <input
                type="number"
                value={lineEndTop}
                onChange={(e) => setLineEndTop(e.target.value)}
                style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
              />
            </div>
          </div>
          <button
            onClick={addLine}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            선 추가
          </button>
        </div>

        {/* 텍스트박스 추가 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #e91e63" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#e91e63" }}>📝 텍스트박스 추가</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={textBoxText}
              onChange={(e) => setTextBoxText(e.target.value)}
              placeholder="텍스트 입력"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <button
            onClick={addTextBox}
            style={{
              padding: "8px 16px",
              backgroundColor: "#e91e63",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            텍스트박스 추가
          </button>
        </div>

        {/* 도형 조작 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 도형 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={shapeName}
              onChange={(e) => setShapeName(e.target.value)}
              placeholder="도형 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getShapeInfo}
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
              onClick={deleteShape}
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
          {result || "위 버튼을 클릭하여 Excel 도형 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Shapes;
