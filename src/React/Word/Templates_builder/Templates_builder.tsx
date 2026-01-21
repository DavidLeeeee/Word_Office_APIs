import React, { useState, useEffect } from "react";

/* global Word */

interface TemplateData {
  name: string;
  contentControls: Array<{
    title: string;
    tag: string;
    placeholderText: string;
    appearance: string;
    color: string;
    position: number; // 문서 내 순서
  }>;
  createdAt: string;
}

const Templates_builder: React.FC = () => {
  const [result, setResult] = useState("");
  const [templateVariable, setTemplateVariable] = useState("");
  const [templateValue, setTemplateValue] = useState("");
  const [templateName, setTemplateName] = useState("");
  const [savedTemplates, setSavedTemplates] = useState<TemplateData[]>([]);

  // 1. Content Control로 템플릿 변수 생성
  const createContentControl = async () => {
    if (!templateVariable.trim()) {
      setResult("템플릿 변수 이름을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (selection.text.trim() === "") {
          setResult("텍스트를 선택한 후 Content Control을 생성해주세요.");
          return;
        }

        // Content Control 생성
        const contentControl = selection.insertContentControl();
        contentControl.title = templateVariable;
        contentControl.tag = `template_${templateVariable}`;
        contentControl.appearance = "BoundingBox"; // 테두리 표시
        contentControl.color = "#4472C4"; // 파란색
        contentControl.placeholderText = `{{${templateVariable}}}`;
        
        await context.sync();

        setResult(`Content Control 생성 완료!\n변수명: "${templateVariable}"\n위치: 선택한 텍스트 영역\n\n과정:\n1. context.document.getSelection()으로 사용자 선택 가져오기\n2. selection.insertContentControl()으로 Content Control 생성\n3. contentControl.title, tag, appearance, color, placeholderText 설정\n4. context.sync()로 동기화\n\n참고: Content Control은 구조화된 입력 필드로, 사용자가 직접 편집할 수 있는 영역입니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 모든 Content Control 목록 보기
  const listContentControls = async () => {
    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("title,tag,text,appearance,color");
        await context.sync();

        if (contentControls.items.length === 0) {
          setResult("Content Control이 없습니다.\n\n과정:\n1. context.document.contentControls로 모든 Content Control 가져오기\n2. contentControls.load()로 필요한 속성 로드\n3. context.sync()로 동기화");
          return;
        }

        const controlList = contentControls.items.map((control, idx) => {
          return `${idx + 1}. "${control.title || "(제목 없음)"}" (태그: ${control.tag || "(태그 없음)"})\n   텍스트: "${control.text}"\n   외관: ${control.appearance}, 색상: ${control.color || "(기본)"}`;
        }).join("\n\n");

        setResult(`Content Control 목록 (${contentControls.items.length}개):\n\n${controlList}\n\n과정:\n1. context.document.contentControls로 모든 Content Control 가져오기\n2. contentControls.load("title,tag,text,appearance,color")로 속성 로드\n3. context.sync()로 동기화\n4. items 배열을 순회하여 정보 표시`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. Content Control에 값 채우기 (제목으로 찾기)
  const fillContentControl = async () => {
    if (!templateVariable.trim() || !templateValue.trim()) {
      setResult("템플릿 변수 이름과 값을 모두 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("title,text");
        await context.sync();

        // 제목으로 Content Control 찾기
        const targetControl = contentControls.items.find(control => control.title === templateVariable);

        if (!targetControl) {
          setResult(`"${templateVariable}"라는 이름의 Content Control을 찾을 수 없습니다.\n\n과정:\n1. context.document.contentControls로 모든 Content Control 가져오기\n2. contentControls.load("title,text")로 속성 로드\n3. context.sync()로 동기화\n4. items.find()로 제목이 일치하는 Control 찾기`);
          return;
        }

        // Content Control의 텍스트 교체
        const range = targetControl.getRange();
        range.insertText(templateValue, Word.InsertLocation.replace);
        await context.sync();

        setResult(`Content Control 값 채우기 완료!\n변수명: "${templateVariable}"\n값: "${templateValue}"\n\n과정:\n1. context.document.contentControls로 모든 Content Control 가져오기\n2. contentControls.load("title,text")로 속성 로드\n3. context.sync()로 동기화\n4. items.find()로 제목이 일치하는 Control 찾기\n5. control.getRange()로 Range 가져오기\n6. range.insertText(value, Word.InsertLocation.replace)로 텍스트 교체\n7. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 템플릿 변수 검색 및 일괄 치환 ({{변수명}} 형식)
  const replaceTemplateVariables = async () => {
    if (!templateVariable.trim() || !templateValue.trim()) {
      setResult("템플릿 변수 이름과 값을 모두 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const searchPattern = `{{${templateVariable}}}`;
        const searchResults = context.document.body.search(searchPattern, {
          matchCase: false,
          matchWholeWord: false,
        });
        searchResults.load("text");
        await context.sync();

        if (searchResults.items.length === 0) {
          setResult(`"${searchPattern}" 패턴을 찾을 수 없습니다.\n\n과정:\n1. context.document.body.search("{{" + templateVariable + "}}")로 템플릿 변수 검색\n2. searchResults.load("text")로 텍스트 속성 로드\n3. context.sync()로 동기화`);
          return;
        }

        // 역순으로 처리하여 인덱스 변경 문제 방지
        for (let i = searchResults.items.length - 1; i >= 0; i--) {
          searchResults.items[i].insertText(templateValue, Word.InsertLocation.replace);
        }
        await context.sync();

        setResult(`템플릿 변수 일괄 치환 완료! (${searchResults.items.length}개)\n변수명: "${templateVariable}"\n값: "${templateValue}"\n\n과정:\n1. context.document.body.search("{{" + templateVariable + "}}")로 템플릿 변수 검색\n2. searchResults.load("text")로 텍스트 속성 로드\n3. context.sync()로 동기화\n4. 역순으로 순회하며 각 결과에 insertText()로 값 교체\n5. context.sync()로 동기화\n\n참고: {{변수명}} 형식의 텍스트를 검색하여 일괄 치환합니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 템플릿 저장 (현재 문서의 Content Control 구조 저장)
  const saveTemplate = async () => {
    if (!templateName.trim()) {
      setResult("템플릿 이름을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("title,tag,placeholderText,appearance,color");
        await context.sync();

        if (contentControls.items.length === 0) {
          setResult("저장할 Content Control이 없습니다. 먼저 Content Control을 생성해주세요.");
          return;
        }

        // Content Control 정보 수집
        const templateData: TemplateData = {
          name: templateName,
          contentControls: contentControls.items.map((control, index) => ({
            title: control.title || "",
            tag: control.tag || "",
            placeholderText: control.placeholderText || "",
            appearance: control.appearance || "BoundingBox",
            color: control.color || "#4472C4",
            position: index,
          })),
          createdAt: new Date().toISOString(),
        };

        // 로컬 스토리지에 저장
        const existingTemplates = JSON.parse(localStorage.getItem("wordTemplates") || "[]");
        existingTemplates.push(templateData);
        localStorage.setItem("wordTemplates", JSON.stringify(existingTemplates));

        setSavedTemplates(existingTemplates);

        setResult(`템플릿 저장 완료!\n템플릿명: "${templateName}"\nContent Control 개수: ${contentControls.items.length}개\n\n과정:\n1. context.document.contentControls로 모든 Content Control 가져오기\n2. contentControls.load()로 속성 로드\n3. context.sync()로 동기화\n4. 각 Content Control의 정보를 수집하여 JSON 형식으로 변환\n5. localStorage.setItem()으로 저장\n\n참고: 저장된 템플릿은 "템플릿 불러오기" 섹션에서 확인할 수 있습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. 저장된 템플릿 목록 불러오기
  const loadSavedTemplates = () => {
    try {
      const templates = JSON.parse(localStorage.getItem("wordTemplates") || "[]");
      setSavedTemplates(templates);
    } catch (error) {
      console.error("템플릿 로드 오류:", error);
    }
  };

  // 7. 템플릿 불러오기 (저장된 템플릿을 현재 문서에 적용)
  const applyTemplate = async (template: TemplateData) => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        
        // 각 Content Control 생성
        for (const controlData of template.contentControls) {
          // 고유한 마커 생성 (타임스탬프 + 랜덤)
          const marker = `__TEMPLATE_MARKER_${Date.now()}_${Math.random().toString(36).substr(2, 9)}__`;
          const placeholder = controlData.placeholderText || `{{${controlData.title}}}`;
          
          // 마커 + 플레이스홀더 텍스트 삽입
          const endRange = body.getRange("End");
          endRange.insertText(marker + placeholder, Word.InsertLocation.before);
          await context.sync();

          // 마커를 기준으로 삽입된 텍스트 범위 찾기
          const searchResults = body.search(marker + placeholder, {
            matchCase: true,
            matchWholeWord: false,
          });
          searchResults.load("text");
          await context.sync();

          if (searchResults.items.length > 0) {
            // 마지막 검색 결과(방금 삽입한 것)에 Content Control 생성
            const insertedRange = searchResults.items[searchResults.items.length - 1];
            
            // 마커 제거하고 플레이스홀더만 남기기
            insertedRange.insertText(placeholder, Word.InsertLocation.replace);
            await context.sync();
            
            // 다시 플레이스홀더로 검색하여 정확한 범위 찾기
            const placeholderResults = body.search(placeholder, {
              matchCase: true,
              matchWholeWord: false,
            });
            placeholderResults.load("text");
            await context.sync();
            
            // 마지막 결과가 방금 삽입한 것
            if (placeholderResults.items.length > 0) {
              const lastPlaceholder = placeholderResults.items[placeholderResults.items.length - 1];
              const contentControl = lastPlaceholder.insertContentControl();
              contentControl.title = controlData.title;
              contentControl.tag = controlData.tag;
              contentControl.appearance = controlData.appearance as Word.ContentControlAppearance;
              contentControl.color = controlData.color;
              contentControl.placeholderText = controlData.placeholderText;
              await context.sync();
            }
          }
          
          // 줄바꿈 추가
          const finalRange = body.getRange("End");
          finalRange.insertText("\n", Word.InsertLocation.before);
          await context.sync();
        }

        setResult(`템플릿 적용 완료!\n템플릿명: "${template.name}"\nContent Control 개수: ${template.contentControls.length}개\n\n과정:\n1. context.document.body.getRange("End")로 문서 끝 위치 가져오기\n2. 저장된 템플릿의 각 Content Control 정보를 순회\n3. 고유한 마커와 함께 플레이스홀더 텍스트 삽입\n4. 마커를 기준으로 삽입된 텍스트 범위 찾기\n5. 마커 제거 후 플레이스홀더만 남기기\n6. 플레이스홀더 범위에 insertContentControl()로 Content Control 생성\n7. context.sync()로 동기화\n\n참고: 템플릿이 문서 끝에 추가됩니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 템플릿 삭제
  const deleteTemplate = (templateNameToDelete: string) => {
    try {
      const templates = JSON.parse(localStorage.getItem("wordTemplates") || "[]");
      const filteredTemplates = templates.filter((t: TemplateData) => t.name !== templateNameToDelete);
      localStorage.setItem("wordTemplates", JSON.stringify(filteredTemplates));
      setSavedTemplates(filteredTemplates);

      setResult(`템플릿 삭제 완료!\n템플릿명: "${templateNameToDelete}"\n\n과정:\n1. localStorage.getItem("wordTemplates")로 저장된 템플릿 가져오기\n2. filter()로 삭제할 템플릿 제외\n3. localStorage.setItem()로 업데이트된 목록 저장`);
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 컴포넌트 마운트 시 저장된 템플릿 목록 불러오기
  useEffect(() => {
    loadSavedTemplates();
  }, []);

  // 9. Content Control 삭제 (제목으로 찾기)
  const deleteContentControl = async () => {
    if (!templateVariable.trim()) {
      setResult("삭제할 Content Control의 변수 이름을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("title");
        await context.sync();

        // 제목으로 Content Control 찾기
        const targetControl = contentControls.items.find(control => control.title === templateVariable);

        if (!targetControl) {
          setResult(`"${templateVariable}"라는 이름의 Content Control을 찾을 수 없습니다.`);
          return;
        }

        // Content Control 삭제 - 범위를 가져와서 텍스트로 교체
        const controlRange = targetControl.getRange();
        controlRange.clear();
        await context.sync();

        setResult(`Content Control 삭제 완료!\n변수명: "${templateVariable}"\n\n과정:\n1. context.document.contentControls로 모든 Content Control 가져오기\n2. contentControls.load("title")로 속성 로드\n3. context.sync()로 동기화\n4. items.find()로 제목이 일치하는 Control 찾기\n5. control.getRange().clear()로 Content Control 내용 삭제\n6. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };


  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word 템플릿 빌더</h3>
        
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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📋 템플릿 빌더란?</h4>
          <p style={{ margin: "0 0 8px 0", color: "#424242" }}>
            템플릿 빌더는 Word 문서에서 <strong>구조화된 입력 필드(Content Control)</strong>를 생성하고 관리하는 도구입니다.
          </p>
          
          <p style={{ margin: "8px 0", color: "#424242", fontWeight: "bold" }}>
            💡 언제 사용하나요?
          </p>
          <div style={{ marginBottom: "10px", padding: "10px", backgroundColor: "#fff", borderRadius: "3px" }}>
            <p style={{ margin: "0 0 5px 0", color: "#424242" }}>
              <strong>예시 1: 반복적인 문서 작성</strong>
            </p>
            <p style={{ margin: "0 0 10px 0", color: "#666", fontSize: "12px" }}>
              "지출결의서", "계약서", "제안서" 등 매번 비슷한 형식의 문서를 작성할 때,<br/>
              문서에 Content Control을 만들어두면 나중에 프로그래밍 방식으로 자동으로 값을 채울 수 있습니다.
            </p>
            
            <p style={{ margin: "10px 0 5px 0", color: "#424242" }}>
              <strong>예시 2: 자동화된 문서 생성</strong>
            </p>
            <p style={{ margin: "0 0 10px 0", color: "#666", fontSize: "12px" }}>
              데이터베이스나 API에서 받은 정보를 Word 문서에 자동으로 채워넣을 때,<br/>
              Content Control을 사용하면 정확한 위치에 값을 넣을 수 있어 오류가 적습니다.
            </p>
            
            <p style={{ margin: "10px 0 5px 0", color: "#424242" }}>
              <strong>예시 3: 사용자 입력 가이드</strong>
            </p>
            <p style={{ margin: "0", color: "#666", fontSize: "12px" }}>
              문서를 다른 사람에게 전달할 때, 특정 위치에만 입력하도록 Content Control을 설정하면<br/>
              사용자가 실수로 다른 부분을 수정하는 것을 방지할 수 있습니다.
            </p>
          </div>
          
          <p style={{ margin: "8px 0", color: "#424242", fontWeight: "bold" }}>
            🔄 텍스트 치환과의 차이점:
          </p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#424242" }}>
            <li><strong>텍스트 치환</strong>: 단순히 {"{{변수명}}"}" 텍스트를 찾아서 바꿈 (위치를 정확히 알 수 없음)</li>
            <li><strong>Content Control</strong>: 문서 내 특정 위치에 구조화된 필드를 생성 (정확한 위치 보장, 재사용 가능)</li>
          </ul>
          
          <p style={{ margin: "8px 0 0 0", color: "#2e7d32", fontSize: "12px", fontWeight: "bold" }}>
            ✅ 템플릿 저장/불러오기 기능이 추가되었습니다! 아래 "템플릿 저장/불러오기" 섹션을 확인하세요.
          </p>
        </div>
        
        {/* 템플릿 변수 생성 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ddd" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>템플릿 변수 생성</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
            <input
              type="text"
              value={templateVariable}
              onChange={(e) => setTemplateVariable(e.target.value)}
              placeholder="변수 이름 (예: 이름, 날짜)"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "200px" }}
            />
            <button
              onClick={createContentControl}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              Content Control 생성
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666" }}>
            사용법: Word 문서에서 텍스트를 선택한 후 변수 이름을 입력하고 버튼을 클릭하세요.
          </div>
        </div>

        {/* 템플릿 변수 목록 보기 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ddd" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>템플릿 변수 목록</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <button
              onClick={listContentControls}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              Content Control 목록
            </button>
          </div>
        </div>

        {/* 템플릿 변수 값 채우기 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ddd" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>템플릿 변수 값 채우기</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
            <input
              type="text"
              value={templateValue}
              onChange={(e) => setTemplateValue(e.target.value)}
              placeholder="값 입력"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "200px" }}
            />
            <button
              onClick={fillContentControl}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              Content Control 채우기
            </button>
            <button
              onClick={replaceTemplateVariables}
              style={{
                padding: "8px 16px",
                backgroundColor: "#e91e63",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              {"{{변수명}}"} 일괄 치환
            </button>
          </div>
          <div style={{ fontSize: "12px", color: "#666" }}>
            변수 이름과 값을 입력한 후 해당 버튼을 클릭하세요. "일괄 치환"은 {"{{변수명}}"} 형식의 텍스트를 모두 찾아 치환합니다.
          </div>
        </div>

        {/* 템플릿 저장/불러오기 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "2px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>💾 템플릿 저장/불러오기</h4>
          
          {/* 차이점 설명 */}
          <div style={{ marginBottom: "15px", padding: "10px", backgroundColor: "#fff3cd", borderRadius: "3px", border: "1px solid #ffc107" }}>
            <h5 style={{ margin: "0 0 8px 0", fontSize: "13px", color: "#856404", fontWeight: "bold" }}>📌 Content Control 생성 vs 템플릿 저장의 차이</h5>
            <div style={{ fontSize: "12px", color: "#856404", lineHeight: "1.6" }}>
              <p style={{ margin: "0 0 5px 0" }}>
                <strong>Content Control 생성:</strong> Word 문서의 <strong>특정 위치</strong>에 입력 필드를 만듭니다. (예: "지출"이라는 텍스트를 선택하고 Content Control로 변환)
              </p>
              <p style={{ margin: "0" }}>
                <strong>템플릿 저장:</strong> 현재 문서에 있는 <strong>모든 Content Control의 구조</strong>를 저장합니다. 나중에 다른 문서에서 이 구조를 불러와 재사용할 수 있습니다. (예: "지출결의서" 템플릿 저장 → 새 문서에서 불러오기)
              </p>
            </div>
          </div>
          
          {/* 템플릿 저장 */}
          <div style={{ marginBottom: "15px", padding: "10px", backgroundColor: "#f1f8f4", borderRadius: "3px" }}>
            <h5 style={{ margin: "0 0 8px 0", fontSize: "13px", color: "#2e7d32" }}>템플릿 저장</h5>
            <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
              <input
                type="text"
                value={templateName}
                onChange={(e) => setTemplateName(e.target.value)}
                placeholder="템플릿 이름 (예: 지출결의서)"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "200px" }}
              />
              <button
                onClick={saveTemplate}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#4caf50",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                템플릿 저장
              </button>
            </div>
            <div style={{ fontSize: "11px", color: "#666", marginTop: "5px" }}>
              현재 문서의 모든 Content Control을 템플릿으로 저장합니다.
            </div>
          </div>

          {/* 저장된 템플릿 목록 */}
          <div style={{ padding: "10px", backgroundColor: "#f1f8f4", borderRadius: "3px" }}>
            <h5 style={{ margin: "0 0 8px 0", fontSize: "13px", color: "#2e7d32" }}>저장된 템플릿 목록</h5>
            {savedTemplates.length === 0 ? (
              <div style={{ fontSize: "12px", color: "#666", fontStyle: "italic" }}>
                저장된 템플릿이 없습니다. 위에서 템플릿을 저장해주세요.
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                {savedTemplates.map((template, idx) => (
                  <div
                    key={idx}
                    style={{
                      padding: "10px",
                      backgroundColor: "#fff",
                      borderRadius: "3px",
                      border: "1px solid #ddd",
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                    }}
                  >
                    <div>
                      <div style={{ fontWeight: "bold", marginBottom: "4px" }}>{template.name}</div>
                      <div style={{ fontSize: "11px", color: "#666" }}>
                        Content Control: {template.contentControls.length}개 | 
                        생성일: {new Date(template.createdAt).toLocaleDateString()}
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: "5px" }}>
                      <button
                        onClick={() => applyTemplate(template)}
                        style={{
                          padding: "6px 12px",
                          backgroundColor: "#2196f3",
                          color: "#fff",
                          border: "none",
                          borderRadius: "3px",
                          cursor: "pointer",
                          fontSize: "12px",
                        }}
                      >
                        불러오기
                      </button>
                      <button
                        onClick={() => deleteTemplate(template.name)}
                        style={{
                          padding: "6px 12px",
                          backgroundColor: "#f44336",
                          color: "#fff",
                          border: "none",
                          borderRadius: "3px",
                          cursor: "pointer",
                          fontSize: "12px",
                        }}
                      >
                        삭제
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* 템플릿 변수 삭제 섹션 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ddd" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#f44336" }}>템플릿 변수 삭제</h4>
          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
            <input
              type="text"
              value={templateVariable}
              onChange={(e) => setTemplateVariable(e.target.value)}
              placeholder="변수 이름"
              style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "200px" }}
            />
            <button
              onClick={deleteContentControl}
              style={{
                padding: "8px 16px",
                backgroundColor: "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              Content Control 삭제
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
          {result || "위 버튼을 클릭하여 템플릿 빌더 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Templates_builder;
