import React, { useState } from "react";

/* global Excel */

const Comment: React.FC = () => {
  const [result, setResult] = useState("");
  const [cellAddress, setCellAddress] = useState("A1");
  const [commentContent, setCommentContent] = useState("주석 내용");
  const [commentId, setCommentId] = useState("");
  const [replyContent, setReplyContent] = useState("답글 내용");
  const [useSelection, setUseSelection] = useState(false);

  // 1. 주석 목록 가져오기
  const listComments = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const comments = sheet.comments;
        comments.load("items/id,items/authorName,items/authorEmail,items/content,items/creationDate,items/resolved");
        await context.sync();

        if (comments.items.length === 0) {
          setResult("현재 워크시트에 주석이 없습니다.");
          return;
        }

        let resultText = `주석 목록 (${comments.items.length}개):\n\n`;
        comments.items.forEach((comment, index) => {
          resultText += `${index + 1}. 주석 ID: ${comment.id}\n`;
          resultText += `   작성자: ${comment.authorName} (${comment.authorEmail})\n`;
          resultText += `   내용: ${comment.content.substring(0, 50)}${comment.content.length > 50 ? "..." : ""}\n`;
          resultText += `   생성일: ${comment.creationDate ? comment.creationDate.toLocaleString() : "없음"}\n`;
          resultText += `   해결됨: ${comment.resolved ? "예" : "아니오"}\n\n`;
        });

        resultText += `과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.comments로 주석 컬렉션 가져오기\n3. comments.load("items/id,items/authorName,...")로 속성 로드\n4. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 주석 추가
  const addComment = async () => {
    if (!commentContent.trim()) {
      setResult("주석 내용을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let cellRange: Excel.Range | string;
        
        if (useSelection) {
          cellRange = context.workbook.getSelectedRange();
          cellRange.load("address");
          await context.sync();
          cellRange = `${sheet.name}!${cellRange.address}`;
        } else {
          cellRange = `${sheet.name}!${cellAddress}`;
        }

        const comments = sheet.comments;
        const newComment = comments.add(cellRange, commentContent);
        newComment.load("id,authorName,authorEmail,content,creationDate");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `"${cellAddress}"`;
        setResult(`주석 추가 완료!\n셀: ${cellRange}\n주석 ID: ${newComment.id}\n작성자: ${newComment.authorName} (${newComment.authorEmail})\n내용: ${newComment.content}\n생성일: ${newComment.creationDate ? newComment.creationDate.toLocaleString() : "없음"}\n\n과정:\n1. sheet.comments.add(${method}, "${commentContent}")로 주석 추가\n2. newComment.load()로 속성 로드\n3. context.sync()로 동기화`);
        setCommentContent("주석 내용");
        setCellAddress("A1");
        setUseSelection(false);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 셀 주소를 확인해주세요. 주석은 단일 셀에만 추가할 수 있습니다.`);
    }
  };

  // 3. 주석 정보 읽기
  const getCommentInfo = async () => {
    if (!commentId.trim()) {
      setResult("주석 ID를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const comments = sheet.comments;
        const comment = comments.getItem(commentId);
        
        comment.load("id,authorName,authorEmail,content,creationDate,resolved,contentType");
        const location = comment.getLocation();
        location.load("address");
        await context.sync();

        const info = `주석 정보:\n\nID: ${comment.id}\n작성자: ${comment.authorName} (${comment.authorEmail})\n내용: ${comment.content}\n생성일: ${comment.creationDate ? comment.creationDate.toLocaleString() : "없음"}\n해결됨: ${comment.resolved ? "예" : "아니오"}\n내용 타입: ${comment.contentType}\n위치: ${location.address}\n\n과정:\n1. sheet.comments.getItem("${commentId}")로 주석 가져오기\n2. comment.load()로 속성 로드\n3. comment.getLocation()로 위치 가져오기\n4. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 주석 ID를 확인해주세요.`);
    }
  };

  // 4. 셀의 주석 가져오기
  const getCommentByCell = async () => {
    if (!useSelection && !cellAddress.trim()) {
      setResult("셀 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let cellRange: Excel.Range | string;
        
        if (useSelection) {
          cellRange = context.workbook.getSelectedRange();
          cellRange.load("address");
          await context.sync();
          cellRange = `${sheet.name}!${cellRange.address}`;
        } else {
          cellRange = `${sheet.name}!${cellAddress}`;
        }

        const comments = sheet.comments;
        const comment = comments.getItemByCell(cellRange);
        comment.load("id,authorName,authorEmail,content,creationDate,resolved");
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `"${cellAddress}"`;
        setResult(`셀의 주석 가져오기 완료!\n셀: ${cellRange}\n주석 ID: ${comment.id}\n작성자: ${comment.authorName} (${comment.authorEmail})\n내용: ${comment.content}\n생성일: ${comment.creationDate ? comment.creationDate.toLocaleString() : "없음"}\n해결됨: ${comment.resolved ? "예" : "아니오"}\n\n과정:\n1. sheet.comments.getItemByCell(${method})로 주석 가져오기\n2. comment.load()로 속성 로드\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      if (error.code === "ItemNotFound") {
        setResult(`해당 셀에 주석이 없습니다.\n셀: ${useSelection ? "선택된 셀" : cellAddress}`);
      } else {
        setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
      }
    }
  };

  // 5. 주석 내용 수정
  const updateComment = async () => {
    if (!commentId.trim()) {
      setResult("주석 ID를 입력해주세요.");
      return;
    }
    if (!commentContent.trim()) {
      setResult("주석 내용을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const comments = sheet.comments;
        const comment = comments.getItem(commentId);
        comment.load("id,content");
        await context.sync();

        const oldContent = comment.content;
        comment.content = commentContent;
        await context.sync();

        setResult(`주석 내용 수정 완료!\n주석 ID: ${comment.id}\n이전 내용: ${oldContent}\n새 내용: ${comment.content}\n\n과정:\n1. sheet.comments.getItem("${commentId}")로 주석 가져오기\n2. comment.content = "${commentContent}"로 내용 수정\n3. context.sync()로 동기화`);
        setCommentContent("주석 내용");
        setCommentId("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 주석 해결/해제
  const toggleResolved = async () => {
    if (!commentId.trim()) {
      setResult("주석 ID를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const comments = sheet.comments;
        const comment = comments.getItem(commentId);
        comment.load("id,resolved");
        await context.sync();

        const oldResolved = comment.resolved;
        comment.resolved = !oldResolved;
        await context.sync();

        setResult(`주석 해결 상태 변경 완료!\n주석 ID: ${comment.id}\n이전 상태: ${oldResolved ? "해결됨" : "미해결"}\n새 상태: ${comment.resolved ? "해결됨" : "미해결"}\n\n과정:\n1. sheet.comments.getItem("${commentId}")로 주석 가져오기\n2. comment.resolved = ${!oldResolved}로 상태 변경\n3. context.sync()로 동기화`);
        setCommentId("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 7. 주석 삭제
  const deleteComment = async () => {
    if (!commentId.trim()) {
      setResult("삭제할 주석 ID를 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const comments = sheet.comments;
        const comment = comments.getItem(commentId);
        comment.load("id");
        await context.sync();

        const deletedId = comment.id;
        comment.delete();
        await context.sync();

        setResult(`주석 삭제 완료!\n삭제된 주석 ID: ${deletedId}\n\n과정:\n1. sheet.comments.getItem("${commentId}")로 주석 가져오기\n2. comment.delete()로 주석 삭제\n3. context.sync()로 동기화`);
        setCommentId("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 8. 답글 추가
  const addReply = async () => {
    if (!commentId.trim()) {
      setResult("주석 ID를 입력해주세요.");
      return;
    }
    if (!replyContent.trim()) {
      setResult("답글 내용을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const comments = sheet.comments;
        const comment = comments.getItem(commentId);
        comment.load("id");
        const replies = comment.replies;
        const newReply = replies.add(replyContent);
        newReply.load("id,authorName,authorEmail,content,creationDate");
        await context.sync();

        setResult(`답글 추가 완료!\n주석 ID: ${comment.id}\n답글 ID: ${newReply.id}\n작성자: ${newReply.authorName} (${newReply.authorEmail})\n내용: ${newReply.content}\n생성일: ${newReply.creationDate ? newReply.creationDate.toLocaleString() : "없음"}\n\n과정:\n1. sheet.comments.getItem("${commentId}")로 주석 가져오기\n2. comment.replies.add("${replyContent}")로 답글 추가\n3. newReply.load()로 속성 로드\n4. context.sync()로 동기화`);
        setReplyContent("답글 내용");
        setCommentId("");
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
        setCellAddress(range.address);
        setUseSelection(true);
        setResult(`선택된 범위: ${range.address}\n이제 '선택된 범위 사용' 모드가 활성화되었습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: Excel에서 셀을 선택한 후 다시 시도해주세요.`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 주석</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>💬 Excel 주석 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 주석은 셀에 대한 피드백과 협업을 위한 기능입니다. 주석을 추가하고, 답글을 달고, 해결 상태를 관리할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>주석 목록 조회</li>
            <li>주석 추가</li>
            <li>주석 정보 읽기</li>
            <li>셀의 주석 가져오기</li>
            <li>주석 내용 수정</li>
            <li>주석 해결/해제</li>
            <li>주석 삭제</li>
            <li>답글 추가</li>
          </ul>
        </div>

        {/* 주석 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 주석 목록</h4>
          <button
            onClick={listComments}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            주석 목록 보기
          </button>
        </div>

        {/* 주석 추가 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 주석 추가</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>셀 주소:</label>
            <input
              type="text"
              value={cellAddress}
              onChange={(e) => setCellAddress(e.target.value)}
              placeholder="예: A1"
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
              선택된 셀 사용
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
                marginBottom: "10px",
              }}
            >
              현재 선택된 셀 가져오기
            </button>
          )}
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>주석 내용:</label>
            <textarea
              value={commentContent}
              onChange={(e) => setCommentContent(e.target.value)}
              placeholder="주석 내용 입력"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minHeight: "60px", marginBottom: "10px" }}
            />
          </div>
          <button
            onClick={addComment}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            주석 추가
          </button>
        </div>

        {/* 주석 조작 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 주석 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>주석 ID:</label>
            <input
              type="text"
              value={commentId}
              onChange={(e) => setCommentId(e.target.value)}
              placeholder="주석 ID"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>주석 내용 (수정용):</label>
            <textarea
              value={commentContent}
              onChange={(e) => setCommentContent(e.target.value)}
              placeholder="주석 내용 입력"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minHeight: "60px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={getCommentInfo}
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
              onClick={updateComment}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              내용 수정
            </button>
            <button
              onClick={toggleResolved}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              해결/해제
            </button>
            <button
              onClick={deleteComment}
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
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>답글 내용:</label>
            <textarea
              value={replyContent}
              onChange={(e) => setReplyContent(e.target.value)}
              placeholder="답글 내용 입력"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minHeight: "60px", marginBottom: "10px" }}
            />
          </div>
          <button
            onClick={addReply}
            style={{
              padding: "8px 16px",
              backgroundColor: "#9c27b0",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            답글 추가
          </button>
        </div>

        {/* 셀의 주석 가져오기 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #607d8b" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#607d8b" }}>🔍 셀의 주석 가져오기</h4>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>셀 주소:</label>
            <input
              type="text"
              value={cellAddress}
              onChange={(e) => setCellAddress(e.target.value)}
              placeholder="예: A1"
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
              선택된 셀 사용
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
                marginBottom: "10px",
              }}
            >
              현재 선택된 셀 가져오기
            </button>
          )}
          <button
            onClick={getCommentByCell}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            셀의 주석 가져오기
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
          {result || "위 버튼을 클릭하여 Excel 주석 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Comment;
