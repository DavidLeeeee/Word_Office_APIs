import React, { useState } from "react";

/* global Word */

const Contents: React.FC = () => {
  const [result, setResult] = useState("");
  const [tableRows, setTableRows] = useState("3");
  const [tableCols, setTableCols] = useState("3");
  const [listText, setListText] = useState("");
  const [hyperlinkText, setHyperlinkText] = useState("");
  const [hyperlinkUrl, setHyperlinkUrl] = useState("");

  // 1. 표(Table) 생성
  const createTable = async () => {
    const rows = parseInt(tableRows) || 3;
    const cols = parseInt(tableCols) || 3;

    if (rows < 1 || cols < 1) {
      setResult("행과 열의 개수는 1 이상이어야 합니다.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const range = body.getRange("End");
        
        // 표 삽입
        const table = range.insertTable(rows, cols, Word.InsertLocation.before);
        table.load("rowCount,columnCount");
        await context.sync();

        setResult(`표 생성 완료!\n행: ${rows}개, 열: ${cols}개\n\n과정:\n1. context.document.body.getRange("End")로 문서 끝 위치 가져오기\n2. range.insertTable(rows, cols, Word.InsertLocation.before)로 표 삽입\n3. table.load("rowCount,columnCount")로 속성 로드\n4. context.sync()로 동기화\n\n참고: 표는 문서 끝에 삽입됩니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 표 목록 보기
  const listTables = async () => {
    try {
      await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("rowCount,columnCount");
        await context.sync();

        if (tables.items.length === 0) {
          setResult("표가 없습니다.\n\n과정:\n1. context.document.body.tables로 모든 표 가져오기\n2. tables.load('rowCount,columnCount')로 속성 로드\n3. context.sync()로 동기화");
          return;
        }

        const tableList = tables.items.map((table, idx) => {
          return `${idx + 1}. 표 ${idx + 1} (${table.rowCount}행 × ${table.columnCount}열)`;
        }).join("\n");

        setResult(`표 목록 (${tables.items.length}개):\n\n${tableList}\n\n과정:\n1. context.document.body.tables로 모든 표 가져오기\n2. tables.load('rowCount,columnCount')로 속성 로드\n3. context.sync()로 동기화\n4. items 배열을 순회하여 정보 표시`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 표 셀에 텍스트 입력
  const fillTableCell = async (tableIndex: number, rowIndex: number, colIndex: number, text: string) => {
    try {
      await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("rowCount,columnCount");
        await context.sync();

        if (tables.items.length <= tableIndex) {
          setResult(`표 ${tableIndex + 1}이 존재하지 않습니다.`);
          return;
        }

        const table = tables.items[tableIndex];
        const rows = table.rows;
        rows.load("cells");
        await context.sync();

        if (rows.items.length <= rowIndex) {
          setResult(`행 ${rowIndex + 1}이 존재하지 않습니다.`);
          return;
        }

        const row = rows.items[rowIndex];
        const cells = row.cells;
        cells.load("body");
        await context.sync();

        if (cells.items.length <= colIndex) {
          setResult(`열 ${colIndex + 1}이 존재하지 않습니다.`);
          return;
        }

        const cell = cells.items[colIndex];
        const cellBody = cell.body;
        cellBody.insertText(text, Word.InsertLocation.replace);
        await context.sync();

        setResult(`표 셀에 텍스트 입력 완료!\n표: ${tableIndex + 1}, 행: ${rowIndex + 1}, 열: ${colIndex + 1}\n텍스트: "${text}"\n\n과정:\n1. context.document.body.tables로 모든 표 가져오기\n2. table.rows.items[rowIndex]로 특정 행 가져오기\n3. row.cells.items[colIndex]로 특정 셀 가져오기\n4. cell.body.insertText()로 텍스트 입력\n5. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 번호 매기기 리스트 생성
  const createNumberedList = async () => {
    if (!listText.trim()) {
      setResult("리스트 항목을 입력해주세요. (줄바꿈으로 구분)");
      return;
    }

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const range = body.getRange("End");
        
        const items = listText.split("\n").filter(item => item.trim() !== "");
        
        for (let i = 0; i < items.length; i++) {
          const item = items[i].trim();
          range.insertText(`${i + 1}. ${item}`, Word.InsertLocation.before);
          
          if (i < items.length - 1) {
            range.insertText("\n", Word.InsertLocation.before);
          }
          
          await context.sync();
        }

        // 번호 매기기 리스트 형식 적용
        const paragraphs = body.paragraphs;
        paragraphs.load("text");
        await context.sync();

        // 마지막 삽입된 항목들 찾기
        const insertedParagraphs = paragraphs.items.slice(-items.length);
        for (const paragraph of insertedParagraphs) {
          paragraph.listItem = {
            level: 0,
            listString: "1",
          };
        }
        await context.sync();

        setResult(`번호 매기기 리스트 생성 완료!\n항목 개수: ${items.length}개\n\n과정:\n1. context.document.body.getRange("End")로 문서 끝 위치 가져오기\n2. 입력된 텍스트를 줄바꿈으로 분리\n3. 각 항목을 "번호. 텍스트" 형식으로 삽입\n4. paragraph.listItem으로 번호 매기기 리스트 형식 적용\n5. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 글머리 기호 리스트 생성
  const createBulletedList = async () => {
    if (!listText.trim()) {
      setResult("리스트 항목을 입력해주세요. (줄바꿈으로 구분)");
      return;
    }

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const range = body.getRange("End");
        
        const items = listText.split("\n").filter(item => item.trim() !== "");
        
        for (let i = 0; i < items.length; i++) {
          const item = items[i].trim();
          range.insertText(`• ${item}`, Word.InsertLocation.before);
          
          if (i < items.length - 1) {
            range.insertText("\n", Word.InsertLocation.before);
          }
          
          await context.sync();
        }

        // 글머리 기호 리스트 형식 적용
        const paragraphs = body.paragraphs;
        paragraphs.load("text");
        await context.sync();

        // 마지막 삽입된 항목들 찾기
        const insertedParagraphs = paragraphs.items.slice(-items.length);
        for (const paragraph of insertedParagraphs) {
          paragraph.listItem = {
            level: 0,
            listString: "•",
          };
        }
        await context.sync();

        setResult(`글머리 기호 리스트 생성 완료!\n항목 개수: ${items.length}개\n\n과정:\n1. context.document.body.getRange("End")로 문서 끝 위치 가져오기\n2. 입력된 텍스트를 줄바꿈으로 분리\n3. 각 항목을 "• 텍스트" 형식으로 삽입\n4. paragraph.listItem으로 글머리 기호 리스트 형식 적용\n5. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. 하이퍼링크 생성
  const createHyperlink = async () => {
    if (!hyperlinkText.trim() || !hyperlinkUrl.trim()) {
      setResult("링크 텍스트와 URL을 모두 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        // 선택된 텍스트가 있으면 그 텍스트에 링크 적용, 없으면 새로 삽입
        if (selection.text.trim() === "") {
          selection.insertText(hyperlinkText, Word.InsertLocation.replace);
          await context.sync();
        }

        // 하이퍼링크 삽입
        const hyperlink = selection.insertHyperlink(hyperlinkUrl, hyperlinkText);
        await context.sync();

        setResult(`하이퍼링크 생성 완료!\n텍스트: "${hyperlinkText}"\nURL: "${hyperlinkUrl}"\n\n과정:\n1. context.document.getSelection()으로 사용자 선택 가져오기\n2. 선택된 텍스트가 없으면 insertText()로 텍스트 삽입\n3. selection.insertHyperlink(url, text)로 하이퍼링크 생성\n4. context.sync()로 동기화\n\n참고: 텍스트를 선택한 후 실행하면 선택된 텍스트에 링크가 적용됩니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 7. 하이퍼링크 목록 보기
  const listHyperlinks = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const hyperlinks = body.hyperlinks;
        hyperlinks.load("address,text");
        await context.sync();

        if (hyperlinks.items.length === 0) {
          setResult("하이퍼링크가 없습니다.\n\n과정:\n1. context.document.body.hyperlinks로 모든 하이퍼링크 가져오기\n2. hyperlinks.load('address,text')로 속성 로드\n3. context.sync()로 동기화");
          return;
        }

        const linkList = hyperlinks.items.map((link, idx) => {
          return `${idx + 1}. "${link.text || "(텍스트 없음)"}" → ${link.address || "(주소 없음)"}`;
        }).join("\n");

        setResult(`하이퍼링크 목록 (${hyperlinks.items.length}개):\n\n${linkList}\n\n과정:\n1. context.document.body.hyperlinks로 모든 하이퍼링크 가져오기\n2. hyperlinks.load('address,text')로 속성 로드\n3. context.sync()로 동기화\n4. items 배열을 순회하여 정보 표시`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 표 삭제
  const deleteTable = async (tableIndex: number) => {
    try {
      await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("rowCount");
        await context.sync();

        if (tables.items.length <= tableIndex) {
          setResult(`표 ${tableIndex + 1}이 존재하지 않습니다.`);
          return;
        }

        const table = tables.items[tableIndex];
        table.delete();
        await context.sync();

        setResult(`표 삭제 완료!\n표 번호: ${tableIndex + 1}\n\n과정:\n1. context.document.body.tables로 모든 표 가져오기\n2. tables.items[index]로 특정 표 가져오기\n3. table.delete()로 표 삭제\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word Contents 작업</h3>
        
        {/* 표 작업 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ddd" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>📊 표(Table) 작업</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
            <input
              type="number"
              value={tableRows}
              onChange={(e) => setTableRows(e.target.value)}
              placeholder="행 개수"
              min="1"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "100px" }}
            />
            <span>×</span>
            <input
              type="number"
              value={tableCols}
              onChange={(e) => setTableCols(e.target.value)}
              placeholder="열 개수"
              min="1"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "100px" }}
            />
            <button
              onClick={createTable}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              표 생성
            </button>
            <button
              onClick={listTables}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              표 목록
            </button>
          </div>
        </div>

        {/* 리스트 작업 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ddd" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📝 리스트 작업</h4>
          <div style={{ marginBottom: "10px" }}>
            <textarea
              value={listText}
              onChange={(e) => setListText(e.target.value)}
              placeholder="리스트 항목을 입력하세요 (줄바꿈으로 구분)"
              rows={4}
              style={{ 
                width: "100%", 
                padding: "8px", 
                border: "1px solid #ddd", 
                borderRadius: "5px",
                fontFamily: "inherit",
                resize: "vertical"
              }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={createNumberedList}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              번호 매기기 리스트 생성
            </button>
            <button
              onClick={createBulletedList}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              글머리 기호 리스트 생성
            </button>
          </div>
        </div>

        {/* 하이퍼링크 작업 섹션 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ddd" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#e91e63" }}>🔗 하이퍼링크 작업</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
            <input
              type="text"
              value={hyperlinkText}
              onChange={(e) => setHyperlinkText(e.target.value)}
              placeholder="링크 텍스트"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "200px" }}
            />
            <input
              type="text"
              value={hyperlinkUrl}
              onChange={(e) => setHyperlinkUrl(e.target.value)}
              placeholder="URL (예: https://example.com)"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "300px" }}
            />
            <button
              onClick={createHyperlink}
              style={{
                padding: "8px 16px",
                backgroundColor: "#e91e63",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              하이퍼링크 생성
            </button>
            <button
              onClick={listHyperlinks}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              하이퍼링크 목록
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666" }}>
            사용법: Word 문서에서 텍스트를 선택한 후 링크를 생성하면 선택된 텍스트에 링크가 적용됩니다. 선택하지 않으면 새 텍스트가 삽입됩니다.
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
          {result || "위 버튼을 클릭하여 Contents 작업 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Contents;
