import React, { useState } from "react";

/* global Excel */

const Queries: React.FC = () => {
  const [result, setResult] = useState("");
  const [queryName, setQueryName] = useState("");

  // 1. 쿼리 목록 가져오기
  const listQueries = async () => {
    try {
      await Excel.run(async (context) => {
        const queries = context.workbook.queries;
        queries.load("items/name,items/error,items/loadedTo,items/loadedToDataModel,items/refreshDate,items/rowsLoadedCount");
        await context.sync();

        if (queries.items.length === 0) {
          setResult("현재 워크북에 쿼리가 없습니다.\n\n참고: 쿼리는 Power Query를 통해 생성되며, Excel UI에서 직접 만들 수 있습니다.");
          return;
        }

        let resultText = `쿼리 목록 (${queries.items.length}개):\n\n`;
        queries.items.forEach((query, index) => {
          resultText += `${index + 1}. ${query.name}\n`;
          resultText += `   오류: ${query.error === "None" ? "없음" : query.error}\n`;
          resultText += `   로드 대상: ${query.loadedTo}\n`;
          resultText += `   데이터 모델 로드: ${query.loadedToDataModel ? "예" : "아니오"}\n`;
          resultText += `   마지막 새로고침: ${query.refreshDate ? query.refreshDate.toLocaleString() : "없음"}\n`;
          resultText += `   로드된 행 수: ${query.rowsLoadedCount >= 0 ? query.rowsLoadedCount : "오류 발생"}\n\n`;
        });

        resultText += `과정:\n1. context.workbook.queries로 쿼리 컬렉션 가져오기\n2. queries.load("items/name,items/error,...")로 속성 로드\n3. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 쿼리 정보 읽기
  const getQueryInfo = async () => {
    if (!queryName.trim()) {
      setResult("쿼리 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const queries = context.workbook.queries;
        const query = queries.getItem(queryName);
        
        query.load("name,error,loadedTo,loadedToDataModel,refreshDate,rowsLoadedCount");
        await context.sync();

        const info = `쿼리 정보:\n\n이름: ${query.name}\n오류: ${query.error === "None" ? "없음" : query.error}\n로드 대상: ${query.loadedTo}\n데이터 모델 로드: ${query.loadedToDataModel ? "예" : "아니오"}\n마지막 새로고침: ${query.refreshDate ? query.refreshDate.toLocaleString() : "없음"}\n로드된 행 수: ${query.rowsLoadedCount >= 0 ? query.rowsLoadedCount : "오류 발생"}\n\n과정:\n1. context.workbook.queries.getItem("${queryName}")로 쿼리 가져오기\n2. query.load()로 속성 로드\n3. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 쿼리 이름을 확인해주세요.`);
    }
  };

  // 3. 쿼리 새로고침
  const refreshQuery = async () => {
    if (!queryName.trim()) {
      setResult("새로고침할 쿼리 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const queries = context.workbook.queries;
        const query = queries.getItem(queryName);
        query.load("name");
        await context.sync();

        const refreshedName = query.name;
        query.refresh();
        await context.sync();

        setResult(`쿼리 새로고침 완료!\n쿼리: ${refreshedName}\n\n과정:\n1. context.workbook.queries.getItem("${queryName}")로 쿼리 가져오기\n2. query.refresh()로 쿼리 새로고침\n3. context.sync()로 동기화`);
        setQueryName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 4. 모든 쿼리 새로고침
  const refreshAllQueries = async () => {
    try {
      await Excel.run(async (context) => {
        const queries = context.workbook.queries;
        queries.load("items/name");
        await context.sync();

        const queryNames = queries.items.map(q => q.name);
        queries.refreshAll();
        await context.sync();

        setResult(`모든 쿼리 새로고침 완료!\n새로고침된 쿼리: ${queryNames.length}개\n${queryNames.length > 0 ? queryNames.map((n, i) => `${i + 1}. ${n}`).join("\n") : "(쿼리 없음)"}\n\n과정:\n1. context.workbook.queries로 쿼리 컬렉션 가져오기\n2. queries.refreshAll()로 모든 쿼리 새로고침\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 쿼리</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>🔍 Excel 쿼리 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 쿼리는 Power Query를 통해 생성된 데이터 연결 및 변환 작업입니다. 외부 데이터 소스에서 데이터를 가져와 변환하고 로드할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>쿼리 목록 조회</li>
            <li>쿼리 정보 읽기</li>
            <li>쿼리 새로고침</li>
            <li>모든 쿼리 새로고침</li>
          </ul>
          <p style={{ margin: "8px 0 0 0", color: "#1976d2", fontSize: "12px", fontStyle: "italic" }}>
            참고: 쿼리는 Excel UI의 "데이터" 탭에서 Power Query를 통해 생성할 수 있습니다.
          </p>
        </div>

        {/* 쿼리 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 쿼리 목록</h4>
          <button
            onClick={listQueries}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              marginRight: "10px",
            }}
          >
            쿼리 목록 보기
          </button>
          <button
            onClick={refreshAllQueries}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            모든 쿼리 새로고침
          </button>
        </div>

        {/* 쿼리 조작 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 쿼리 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={queryName}
              onChange={(e) => setQueryName(e.target.value)}
              placeholder="쿼리 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={getQueryInfo}
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
              onClick={refreshQuery}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              새로고침
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
          {result || "위 버튼을 클릭하여 Excel 쿼리 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Queries;
