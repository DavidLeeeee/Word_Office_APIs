import React, { useState } from "react";

/* global Word */

const Comment: React.FC = () => {
  const [result, setResult] = useState("");
  const [commentText, setCommentText] = useState("");
  const [replyText, setReplyText] = useState("");

  // 1. 선택된 텍스트에 주석 추가
  const insertComment = async () => {
    if (!commentText.trim()) {
      setResult("주석 내용을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (selection.text.trim() === "") {
          setResult("주석을 추가할 텍스트를 먼저 선택해주세요.");
          return;
        }

        const comment = selection.insertComment(commentText);
        comment.load("id,authorName,authorEmail,creationDate");
        await context.sync();

        setResult(`주석 추가 완료!\n\n주석 ID: ${comment.id}\n작성자: ${comment.authorName} (${comment.authorEmail})\n작성일: ${comment.creationDate.toLocaleString()}\n내용: ${commentText}\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.insertComment(commentText)로 주석 추가\n3. comment.load()로 주석 정보 로드\n4. context.sync()로 동기화`);
        
        setCommentText(""); // 입력 필드 초기화
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 모든 주석 목록 보기
  const listComments = async () => {
    try {
      await Word.run(async (context) => {
        const comments = context.document.body.getComments();
        comments.load("items/id,items/authorName,items/authorEmail,items/creationDate,items/content,items/resolved");
        await context.sync();

        if (comments.items.length === 0) {
          setResult("문서에 주석이 없습니다.\n\n과정:\n1. context.document.body.getComments()로 모든 주석 가져오기\n2. comments.load()로 주석 정보 로드\n3. context.sync()로 동기화");
          return;
        }

        const commentList = comments.items.map((comment, idx) => {
          const status = comment.resolved ? "✅ 해결됨" : "⏳ 진행중";
          return `${idx + 1}. [${status}] ${comment.authorName} (${comment.authorEmail})\n   작성일: ${comment.creationDate.toLocaleString()}\n   내용: ${comment.content}\n   ID: ${comment.id}`;
        }).join("\n\n");

        setResult(`주석 목록 (${comments.items.length}개):\n\n${commentList}\n\n과정:\n1. context.document.body.getComments()로 모든 주석 가져오기\n2. comments.load()로 주석 정보 로드\n3. context.sync()로 동기화\n4. items 배열을 순회하여 정보 표시`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 선택된 텍스트의 주석에 답변 추가
  const replyToComment = async () => {
    if (!replyText.trim()) {
      setResult("답변 내용을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const comments = selection.getComments();
        comments.load("items/id,items/resolved");
        await context.sync();

        if (comments.items.length === 0) {
          setResult("선택된 텍스트에 주석이 없습니다. 주석이 있는 텍스트를 선택해주세요.");
          return;
        }

        // 첫 번째 주석에 답변 추가
        const firstComment = comments.items[0];
        if (firstComment.resolved) {
          setResult("이미 해결된 주석에는 답변을 추가할 수 없습니다.");
          return;
        }

        const reply = firstComment.reply(replyText);
        reply.load("id,authorName,authorEmail,creationDate");
        await context.sync();

        setResult(`답변 추가 완료!\n\n원본 주석 ID: ${firstComment.id}\n답변 ID: ${reply.id}\n작성자: ${reply.authorName} (${reply.authorEmail})\n작성일: ${reply.creationDate.toLocaleString()}\n내용: ${replyText}\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.getComments()로 해당 텍스트의 주석 가져오기\n3. comment.reply(replyText)로 답변 추가\n4. reply.load()로 답변 정보 로드\n5. context.sync()로 동기화`);
        
        setReplyText(""); // 입력 필드 초기화
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 주석 해결 처리 (resolved = true)
  const resolveComment = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const comments = selection.getComments();
        comments.load("items/id,items/resolved");
        await context.sync();

        if (comments.items.length === 0) {
          setResult("선택된 텍스트에 주석이 없습니다. 주석이 있는 텍스트를 선택해주세요.");
          return;
        }

        const firstComment = comments.items[0];
        if (firstComment.resolved) {
          setResult("이미 해결된 주석입니다.");
          return;
        }

        firstComment.resolved = true;
        await context.sync();

        setResult(`주석 해결 처리 완료!\n\n주석 ID: ${firstComment.id}\n상태: 해결됨\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.getComments()로 해당 텍스트의 주석 가져오기\n3. comment.resolved = true로 해결 처리\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 주석 삭제
  const deleteComment = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const comments = selection.getComments();
        comments.load("items/id");
        await context.sync();

        if (comments.items.length === 0) {
          setResult("선택된 텍스트에 주석이 없습니다. 주석이 있는 텍스트를 선택해주세요.");
          return;
        }

        const firstComment = comments.items[0];
        const commentId = firstComment.id;
        firstComment.delete();
        await context.sync();

        setResult(`주석 삭제 완료!\n\n삭제된 주석 ID: ${commentId}\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.getComments()로 해당 텍스트의 주석 가져오기\n3. comment.delete()로 주석 삭제 (답변도 함께 삭제됨)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. 변경 추적 모드 확인 및 설정
  const checkTrackChanges = async () => {
    try {
      await Word.run(async (context) => {
        const document = context.document;
        document.load("changeTrackingMode");
        await context.sync();

        const mode = document.changeTrackingMode;
        const modeText = mode === Word.ChangeTrackingMode.Off ? "비활성화" : 
                        mode === Word.ChangeTrackingMode.TrackAll ? "모든 변경 추적" : 
                        mode === Word.ChangeTrackingMode.TrackMoves ? "이동 추적" : "알 수 없음";

        setResult(`변경 추적 모드: ${modeText}\n\n현재 상태: ${mode === Word.ChangeTrackingMode.Off ? "변경 추적이 비활성화되어 있습니다." : "변경 추적이 활성화되어 있습니다."}\n\n과정:\n1. context.document.load("changeTrackingMode")로 변경 추적 모드 로드\n2. context.sync()로 동기화\n3. document.changeTrackingMode로 모드 확인\n\n참고: Word JavaScript API에서는 변경 추적 모드를 읽을 수만 있고, 설정은 Word UI에서 해야 합니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 7. 검토된 텍스트 가져오기 (변경 추적이 활성화된 경우)
  const getReviewedText = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        
        // 원본 텍스트
        selection.load("text");
        await context.sync();
        const originalText = selection.text;

        // 검토된 텍스트 (변경 추적 반영)
        const reviewedText = selection.getReviewedText(Word.ChangeTrackingVersion.Original);
        reviewedText.load("text");
        await context.sync();

        setResult(`텍스트 비교:\n\n원본 텍스트:\n${originalText}\n\n검토된 텍스트 (원본 버전):\n${reviewedText.text}\n\n과정:\n1. context.document.getSelection()으로 선택된 텍스트 가져오기\n2. selection.load("text")로 원본 텍스트 로드\n3. selection.getReviewedText(Word.ChangeTrackingVersion.Original)로 검토된 텍스트 가져오기\n4. context.sync()로 동기화\n\n참고: 변경 추적이 활성화되어 있어야 차이가 나타납니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: 변경 추적이 활성화되어 있지 않으면 이 기능을 사용할 수 없습니다.`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word 주석 및 검토 기능</h3>

        {/* 안내 섹션 */}
        <div style={{
          marginBottom: "20px",
          padding: "15px",
          backgroundColor: "#e3f2fd",
          borderRadius: "5px",
          border: "1px solid #90caf9",
          fontSize: "13px",
          lineHeight: "1.6"
        }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>💬 주석 및 검토 기능 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#424242" }}>
            이 섹션에서는 Word 문서의 <strong>주석(Comments)</strong>과 <strong>변경 추적(Track Changes)</strong> 기능을 테스트합니다.
          </p>
          <p style={{ margin: "8px 0", color: "#424242", fontWeight: "bold" }}>
            📌 주요 기능:
          </p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#424242" }}>
            <li><strong>주석 추가:</strong> 문서의 특정 부분에 메모를 추가하여 피드백을 남길 수 있습니다.</li>
            <li><strong>주석 답변:</strong> 다른 사람의 주석에 답변을 달아 대화를 이어갈 수 있습니다.</li>
            <li><strong>주석 해결:</strong> 주석이 처리되면 "해결됨"으로 표시할 수 있습니다.</li>
            <li><strong>변경 추적:</strong> 문서 수정 시 변경사항을 추적하여 누가 무엇을 변경했는지 확인할 수 있습니다.</li>
          </ul>
          <p style={{ margin: "8px 0 0 0", color: "#d32f2f", fontSize: "12px", fontStyle: "italic" }}>
            ⚠️ 참고: Word JavaScript API에서는 변경 추적 모드를 읽을 수만 있고, 설정은 Word UI에서 해야 합니다.
          </p>
        </div>

        {/* 주석 작업 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>💬 주석 작업</h4>
          
          <div style={{ marginBottom: "15px" }}>
            <label style={{ display: "block", marginBottom: "5px", fontSize: "13px", fontWeight: "bold" }}>주석 추가</label>
            <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
              <input
                type="text"
                value={commentText}
                onChange={(e) => setCommentText(e.target.value)}
                placeholder="주석 내용을 입력하세요"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "300px" }}
              />
              <button
                onClick={insertComment}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#4caf50",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                주석 추가
              </button>
            </div>
            <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
              사용법: Word 문서에서 텍스트를 선택한 후 주석 내용을 입력하고 버튼을 클릭하세요.
            </div>
          </div>

          <div style={{ marginBottom: "15px" }}>
            <label style={{ display: "block", marginBottom: "5px", fontSize: "13px", fontWeight: "bold" }}>주석 답변</label>
            <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
              <input
                type="text"
                value={replyText}
                onChange={(e) => setReplyText(e.target.value)}
                placeholder="답변 내용을 입력하세요"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "300px" }}
              />
              <button
                onClick={replyToComment}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#2196f3",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                답변 추가
              </button>
            </div>
            <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
              사용법: 주석이 있는 텍스트를 선택한 후 답변 내용을 입력하고 버튼을 클릭하세요.
            </div>
          </div>

          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={listComments}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              주석 목록
            </button>
            <button
              onClick={resolveComment}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              주석 해결
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
              주석 삭제
            </button>
          </div>
        </div>

        {/* 변경 추적 섹션 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>📝 변경 추적</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={checkTrackChanges}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              변경 추적 모드 확인
            </button>
            <button
              onClick={getReviewedText}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              검토된 텍스트 가져오기
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666", marginTop: "10px" }}>
            참고: 변경 추적 모드는 Word UI에서 활성화/비활성화할 수 있습니다. (검토 탭 → 변경 내용 추적)
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
          {result || "위 버튼을 클릭하여 주석 및 검토 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Comment;
