import React, { useState } from "react";

/* global Excel */

const Slicer: React.FC = () => {
  const [result, setResult] = useState("");
  const [slicerName, setSlicerName] = useState("");
  const [tableName, setTableName] = useState("");
  const [columnName, setColumnName] = useState("");

  // 1. 슬라이서 목록 가져오기
  const listSlicers = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const slicers = sheet.slicers;
        slicers.load("items/name,items/id,items/caption,items/width,items/height");
        await context.sync();

        if (slicers.items.length === 0) {
          setResult("현재 워크시트에 슬라이서가 없습니다.");
          return;
        }

        let resultText = `슬라이서 목록 (${slicers.items.length}개):\n\n`;
        slicers.items.forEach((slicer, index) => {
          resultText += `${index + 1}. ${slicer.name}\n`;
          resultText += `   ID: ${slicer.id}\n`;
          resultText += `   캡션: ${slicer.caption || "(없음)"}\n`;
          resultText += `   크기: ${slicer.width}pt × ${slicer.height}pt\n\n`;
        });

        resultText += `과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.slicers로 슬라이서 컬렉션 가져오기\n3. slicers.load("items/name,items/id,...")로 속성 로드\n4. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 슬라이서 생성
  const createSlicer = async () => {
    if (!tableName.trim() || !columnName.trim()) {
      setResult("테이블 이름과 열 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const tables = sheet.tables;
        const table = tables.getItem(tableName);
        table.load("name");
        const columns = table.columns;
        const column = columns.getItem(columnName);
        column.load("name");
        await context.sync();

        const slicers = sheet.slicers;
        const newSlicer = slicers.add(table, column);
        newSlicer.load("name,id,caption,width,height");
        await context.sync();

        setResult(`슬라이서 생성 완료!\n슬라이서 이름: ${newSlicer.name}\nID: ${newSlicer.id}\n캡션: ${newSlicer.caption || "(없음)"}\n크기: ${newSlicer.width}pt × ${newSlicer.height}pt\n테이블: ${table.name}\n열: ${column.name}\n\n과정:\n1. sheet.tables.getItem("${tableName}")로 테이블 가져오기\n2. table.columns.getItem("${columnName}")로 열 가져오기\n3. sheet.slicers.add(table, column)로 슬라이서 생성\n4. newSlicer.load()로 속성 로드\n5. context.sync()로 동기화`);
        setTableName("");
        setColumnName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 테이블과 열 이름을 확인해주세요.`);
    }
  };

  // 3. 슬라이서 정보 읽기
  const getSlicerInfo = async () => {
    if (!slicerName.trim()) {
      setResult("슬라이서 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const slicers = sheet.slicers;
        const slicer = slicers.getItem(slicerName);
        
        slicer.load("name,id,caption,width,height,left,top,sortBy,style,isFilterCleared");
        await context.sync();

        const info = `슬라이서 정보:\n\n이름: ${slicer.name}\nID: ${slicer.id}\n캡션: ${slicer.caption || "(없음)"}\n위치: (${slicer.left}pt, ${slicer.top}pt)\n크기: ${slicer.width}pt × ${slicer.height}pt\n정렬: ${slicer.sortBy}\n스타일: ${slicer.style}\n필터 지워짐: ${slicer.isFilterCleared ? "예" : "아니오"}\n\n과정:\n1. sheet.slicers.getItem("${slicerName}")로 슬라이서 가져오기\n2. slicer.load()로 속성 로드\n3. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 슬라이서 이름을 확인해주세요.`);
    }
  };

  // 4. 슬라이서 삭제
  const deleteSlicer = async () => {
    if (!slicerName.trim()) {
      setResult("삭제할 슬라이서 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const slicers = sheet.slicers;
        const slicer = slicers.getItem(slicerName);
        slicer.load("name");
        await context.sync();

        const deletedName = slicer.name;
        slicer.delete();
        await context.sync();

        setResult(`슬라이서 삭제 완료!\n삭제된 슬라이서: ${deletedName}\n\n과정:\n1. sheet.slicers.getItem("${slicerName}")로 슬라이서 가져오기\n2. slicer.delete()로 슬라이서 삭제\n3. context.sync()로 동기화`);
        setSlicerName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 슬라이서</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>🔍 Excel 슬라이서 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 슬라이서는 테이블이나 피벗 테이블의 데이터를 필터링하는 시각적 인터페이스입니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>슬라이서 생성 (테이블/피벗 테이블 기반)</li>
            <li>슬라이서 목록 조회</li>
            <li>슬라이서 정보 읽기</li>
            <li>슬라이서 삭제</li>
          </ul>
        </div>

        {/* 슬라이서 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 슬라이서 목록</h4>
          <button
            onClick={listSlicers}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            슬라이서 목록 보기
          </button>
        </div>

        {/* 슬라이서 생성 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 슬라이서 생성</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>테이블 이름:</label>
            <input
              type="text"
              value={tableName}
              onChange={(e) => setTableName(e.target.value)}
              placeholder="예: Table1"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>열 이름:</label>
            <input
              type="text"
              value={columnName}
              onChange={(e) => setColumnName(e.target.value)}
              placeholder="예: 열1"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <button
            onClick={createSlicer}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            슬라이서 생성
          </button>
        </div>

        {/* 슬라이서 조작 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 슬라이서 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={slicerName}
              onChange={(e) => setSlicerName(e.target.value)}
              placeholder="슬라이서 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getSlicerInfo}
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
              onClick={deleteSlicer}
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
          {result || "위 버튼을 클릭하여 Excel 슬라이서 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Slicer;
