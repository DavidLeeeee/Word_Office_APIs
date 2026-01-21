import React, { useState } from "react";

/* global Word */

const Shapes: React.FC = () => {
  const [result, setResult] = useState("");
  const [imageUrl, setImageUrl] = useState("");

  // 1. 이미지 삽입 (Base64)
  const insertImage = async () => {
    if (!imageUrl.trim()) {
      setResult("Base64 이미지 데이터를 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const range = body.getRange("End");
        
        // Base64 이미지 확인 및 처리
        let base64Data = imageUrl.trim();
        
        // data:image 형식인 경우 처리
        if (base64Data.startsWith("data:image")) {
          base64Data = base64Data.replace(/^data:image\/\w+;base64,/, "");
        }
        
        // Base64 데이터 유효성 검사
        if (!base64Data || base64Data.length === 0) {
          setResult("올바른 Base64 이미지 데이터를 입력해주세요.");
          return;
        }
        
        // Base64 이미지 삽입
        const picture = range.insertInlinePictureFromBase64(base64Data, Word.InsertLocation.before);
        picture.load("width,height");
        await context.sync();
        
        setResult(`이미지 삽입 완료! (Base64)\n크기: ${picture.width}pt × ${picture.height}pt\n\n과정:\n1. context.document.body.getRange("End")로 문서 끝 위치 가져오기\n2. Base64 데이터에서 data URL 부분 제거 (있는 경우)\n3. range.insertInlinePictureFromBase64()로 이미지 삽입\n4. picture.load("width,height")로 크기 정보 로드\n5. context.sync()로 동기화\n\n참고: Base64 이미지는 지원되지만, 이미지가 너무 크면 성능에 영향을 줄 수 있습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n가능한 원인:\n1. Base64 데이터가 올바르지 않음\n2. 지원되지 않는 이미지 형식\n3. 이미지 데이터가 너무 큼\n4. Base64 문자열에 잘못된 문자가 포함됨`);
    }
  };

  // 2. URL 이미지 삽입 시도 (오류 확인용)
  const insertImageFromUrl = async () => {
    if (!imageUrl.trim()) {
      setResult("이미지 URL을 입력해주세요.");
      return;
    }

    if (!imageUrl.startsWith("http://") && !imageUrl.startsWith("https://")) {
      setResult("올바른 URL 형식이 아닙니다. http:// 또는 https://로 시작해야 합니다.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const range = body.getRange("End");
        
        // URL을 Base64로 변환 시도 (실제로는 불가능)
        // Word API는 URL을 직접 지원하지 않으므로 오류 발생
        
        // 방법 1: insertInlinePictureFromBase64에 URL을 넣어보기
        try {
          const picture = range.insertInlinePictureFromBase64(imageUrl, Word.InsertLocation.before);
          await context.sync();
          setResult("이미지 삽입 성공 (예상치 못한 결과)");
        } catch (base64Error) {
          throw base64Error;
        }
      });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      const errorStack = error instanceof Error && error.stack ? error.stack : "";
      
      setResult(`❌ URL 이미지 삽입 실패!\n\n오류 메시지:\n${errorMessage}\n\n오류 상세:\n${errorStack || "스택 정보 없음"}\n\n결론:\nWord JavaScript API는 보안상의 이유로 외부 URL에서 직접 이미지를 로드하는 것을 지원하지 않습니다.\n\n해결 방법:\n1. 이미지를 Base64로 변환하여 사용\n2. 서버에서 이미지를 다운로드하여 Base64로 변환 후 삽입\n3. Word UI에서 직접 이미지 삽입 기능 사용\n\n과정:\n1. context.document.body.getRange("End")로 문서 끝 위치 가져오기\n2. range.insertInlinePictureFromBase64(url)로 URL 삽입 시도\n3. Word API가 URL을 지원하지 않아 오류 발생\n4. 오류 메시지 확인`);
    }
  };

  // 3. 모든 인라인 그림 목록 보기
  const listInlinePictures = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const inlinePictures = body.inlinePictures;
        inlinePictures.load("width,height");
        await context.sync();

        if (inlinePictures.items.length === 0) {
          setResult("인라인 그림이 없습니다.\n\n과정:\n1. context.document.body.inlinePictures로 모든 인라인 그림 가져오기\n2. inlinePictures.load('width,height')로 속성 로드\n3. context.sync()로 동기화");
          return;
        }

        const pictureList = inlinePictures.items.map((picture, idx) => {
          return `${idx + 1}. 그림 ${idx + 1} (${picture.width}pt × ${picture.height}pt)`;
        }).join("\n");

        setResult(`인라인 그림 목록 (${inlinePictures.items.length}개):\n\n${pictureList}\n\n과정:\n1. context.document.body.inlinePictures로 모든 인라인 그림 가져오기\n2. inlinePictures.load('width,height')로 속성 로드\n3. context.sync()로 동기화\n4. items 배열을 순회하여 정보 표시`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 인라인 그림 삭제
  const deleteInlinePicture = async (pictureIndex: number) => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const inlinePictures = body.inlinePictures;
        inlinePictures.load("width");
        await context.sync();

        if (inlinePictures.items.length <= pictureIndex) {
          setResult(`그림 ${pictureIndex + 1}이 존재하지 않습니다.`);
          return;
        }

        const picture = inlinePictures.items[pictureIndex];
        picture.delete();
        await context.sync();

        setResult(`인라인 그림 삭제 완료!\n그림 번호: ${pictureIndex + 1}\n\n과정:\n1. context.document.body.inlinePictures로 모든 인라인 그림 가져오기\n2. inlinePictures.items[index]로 특정 그림 가져오기\n3. picture.delete()로 그림 삭제\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word Images</h3>
        
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
          <h4 style={{ margin: "0 0 10px 0", color: "#856404", fontSize: "14px" }}>📌 Word API의 이미지 삽입 제한사항</h4>
          <div style={{ color: "#856404" }}>
            <p style={{ margin: "0 0 8px 0", fontWeight: "bold" }}>✅ Base64 이미지 삽입:</p>
            <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px" }}>
              <li>Base64 형식의 이미지 데이터는 삽입 가능합니다.</li>
              <li>지원 형식: PNG, JPEG, GIF 등</li>
              <li>주의: 이미지가 너무 크면 성능에 영향을 줄 수 있습니다.</li>
            </ul>
            <p style={{ margin: "0 0 8px 0", fontWeight: "bold" }}>❌ URL 이미지 삽입:</p>
            <ul style={{ margin: "0", paddingLeft: "20px" }}>
              <li>보안상의 이유로 외부 URL에서 직접 이미지를 로드하지 않습니다.</li>
              <li>URL 이미지를 사용하려면 먼저 Base64로 변환해야 합니다.</li>
            </ul>
            <p style={{ margin: "8px 0 0 0", fontWeight: "bold" }}>⚠️ 도형(Shapes) 삽입:</p>
            <p style={{ margin: "0", fontSize: "12px" }}>
              Word JavaScript API에서는 도형 삽입이 제한적입니다. 도형은 Word UI에서 직접 삽입하거나, OOXML을 사용한 고급 방법이 필요합니다.
            </p>
          </div>
        </div>

        {/* Base64 이미지 작업 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>✅ Base64 이미지 삽입</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
            <input
              type="text"
              value={imageUrl}
              onChange={(e) => setImageUrl(e.target.value)}
              placeholder="Base64 이미지 데이터 (data:image/... 또는 순수 Base64)"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "300px" }}
            />
            <button
              onClick={insertImage}
              style={{
                padding: "8px 16px",
                backgroundColor: "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              Base64 이미지 삽입
            </button>
            <button
              onClick={listInlinePictures}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              이미지 목록
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666" }}>
            사용법: Base64 형식의 이미지 데이터를 입력하세요.
            <br />
            예시: <code style={{ backgroundColor: "#f0f0f0", padding: "2px 4px", borderRadius: "2px" }}>data:image/png;base64,iVBORw0KGgo...</code> 또는 <code style={{ backgroundColor: "#f0f0f0", padding: "2px 4px", borderRadius: "2px" }}>iVBORw0KGgo...</code>
          </div>
        </div>

        {/* URL 이미지 삽입 시도 섹션 (오류 확인용) */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "2px solid #f44336" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#f44336" }}>❌ URL 이미지 삽입 시도 (오류 확인용)</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
            <input
              type="text"
              value={imageUrl}
              onChange={(e) => setImageUrl(e.target.value)}
              placeholder="이미지 URL (예: https://example.com/image.png)"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "300px" }}
            />
            <button
              onClick={insertImageFromUrl}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              URL 이미지 삽입 시도
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#d32f2f", fontWeight: "bold" }}>
            ⚠️ 이 기능은 Word API의 제한사항을 확인하기 위한 테스트용입니다.
            <br />
            URL을 입력하고 버튼을 클릭하면 실제 오류 메시지를 확인할 수 있습니다.
            <br />
            <span style={{ color: "#666", fontWeight: "normal" }}>참고: Word API는 보안상의 이유로 외부 URL에서 직접 이미지를 로드하지 않습니다.</span>
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
          {result || "위 버튼을 클릭하여 Shapes & Images 작업 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Shapes;
