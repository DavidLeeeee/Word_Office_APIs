import React, { useState } from "react";

/* global Word */

const Metadata: React.FC = () => {
  const [result, setResult] = useState("");
  const [customPropName, setCustomPropName] = useState("");
  const [customPropValue, setCustomPropValue] = useState("");
  const [xmlContent, setXmlContent] = useState("");
  const [settingKey, setSettingKey] = useState("");
  const [settingValue, setSettingValue] = useState("");

  // 1. 기본 문서 속성 읽기
  const readBuiltInProperties = async () => {
    try {
      await Word.run(async (context) => {
        const properties = context.document.properties;
        properties.load("title,author,subject,keywords,category,comments,company,manager");
        await context.sync();

        setResult(`기본 문서 속성:\n\n제목: ${properties.title || "(없음)"}\n작성자: ${properties.author || "(없음)"}\n주제: ${properties.subject || "(없음)"}\n키워드: ${properties.keywords || "(없음)"}\n카테고리: ${properties.category || "(없음)"}\n설명: ${properties.comments || "(없음)"}\n회사: ${properties.company || "(없음)"}\n관리자: ${properties.manager || "(없음)"}\n\n과정:\n1. context.document.properties로 문서 속성 가져오기\n2. properties.load()로 필요한 속성 로드\n3. context.sync()로 동기화\n\n확인 방법:\n- Word UI: 파일 > 정보 > 속성 > 고급 속성\n- 또는 파일 탐색기에서 파일 우클릭 > 속성 > 세부 정보`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: 일부 Word 버전에서는 properties API가 제한적일 수 있습니다.`);
    }
  };

  // 2. 기본 문서 속성 설정
  const setBuiltInProperty = async (propName: string, propValue: string) => {
    if (!propValue.trim()) {
      setResult(`${propName} 값을 입력해주세요.`);
      return;
    }

    try {
      await Word.run(async (context) => {
        const properties = context.document.properties;
        
        // properties 객체를 먼저 로드 (필요한 경우)
        properties.load(propName);
        await context.sync();
        
        // 각 속성별로 직접 설정 (동적 속성 접근은 지원되지 않음)
        if (propName === "title") {
          properties.title = propValue;
        } else if (propName === "author") {
          properties.author = propValue;
        } else if (propName === "subject") {
          properties.subject = propValue;
        } else if (propName === "keywords") {
          properties.keywords = propValue;
        } else if (propName === "category") {
          properties.category = propValue;
        } else if (propName === "comments") {
          properties.comments = propValue;
        } else if (propName === "company") {
          properties.company = propValue;
        } else if (propName === "manager") {
          properties.manager = propValue;
        } else {
          setResult(`지원하지 않는 속성입니다: ${propName}`);
          return;
        }
        
        await context.sync();

        setResult(`${propName} 설정 완료!\n값: ${propValue}\n\n과정:\n1. context.document.properties로 문서 속성 가져오기\n2. properties.load("${propName}")로 속성 로드\n3. properties.${propName} = "${propValue}"로 속성 설정\n4. context.sync()로 동기화\n\n확인 방법: Word UI에서 파일 > 정보 > 속성에서 확인 가능`);
      });
    } catch (error: any) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      const errorCode = error?.code || "알 수 없음";
      const errorName = error?.name || "알 수 없음";
      
      setResult(`오류 발생!\n\n오류 이름: ${errorName}\n오류 코드: ${errorCode}\n오류 메시지: ${errorMessage}\n\n가능한 원인:\n1. Word JavaScript API에서 해당 속성 설정이 지원되지 않음\n2. Word 버전이 낮아서 기능이 지원되지 않음\n3. Office Online에서는 일부 속성 설정이 제한될 수 있음\n\n참고: Word JavaScript API의 document.properties는 읽기는 가능하지만, 쓰기는 제한적일 수 있습니다.`);
    }
  };

  // 3. 커스텀 속성 추가
  const addCustomProperty = async () => {
    if (!customPropName.trim() || !customPropValue.trim()) {
      setResult("커스텀 속성 이름과 값을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const properties = context.document.properties;
        const customProps = properties.customProperties;
        customProps.add(customPropName, customPropValue);
        await context.sync();

        setResult(`커스텀 속성 추가 완료!\n이름: ${customPropName}\n값: ${customPropValue}\n\n과정:\n1. context.document.properties.customProperties로 커스텀 속성 컬렉션 가져오기\n2. customProperties.add(name, value)로 속성 추가\n3. context.sync()로 동기화\n\n확인 방법:\n- Word UI: 파일 > 정보 > 속성 > 고급 속성 > 사용자 지정 탭\n- 또는 파일 탐색기에서 파일 우클릭 > 속성 > 세부 정보`);
        
        setCustomPropName("");
        setCustomPropValue("");
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}\n\n참고: 웹 버전(Office Online)에서는 커스텀 속성 추가가 제한적일 수 있습니다.`);
    }
  };

  // 4. 모든 커스텀 속성 읽기
  const readCustomProperties = async () => {
    try {
      await Word.run(async (context) => {
        const properties = context.document.properties;
        const customProps = properties.customProperties;
        customProps.load("items/key,items/value");
        await context.sync();

        if (customProps.items.length === 0) {
          setResult("커스텀 속성이 없습니다.\n\n과정:\n1. context.document.properties.customProperties로 커스텀 속성 컬렉션 가져오기\n2. customProperties.load()로 속성 로드\n3. context.sync()로 동기화");
          return;
        }

        const propList = customProps.items.map((prop: any, idx: number) => {
          return `${idx + 1}. ${prop.key}: ${prop.value}`;
        }).join("\n");

        setResult(`커스텀 속성 목록 (${customProps.items.length}개):\n\n${propList}\n\n과정:\n1. context.document.properties.customProperties로 커스텀 속성 컬렉션 가져오기\n2. customProperties.load()로 속성 로드\n3. context.sync()로 동기화\n4. items 배열을 순회하여 정보 표시`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 커스텀 XML 파트 추가
  const addCustomXmlPart = async () => {
    if (!xmlContent.trim()) {
      setResult("XML 내용을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const xmlParts = context.document.customXmlParts;
        const xmlPart = xmlParts.add(xmlContent);
        xmlPart.load("id,namespaceUri");
        await context.sync();

        setResult(`커스텀 XML 파트 추가 완료!\n\nID: ${xmlPart.id}\nNamespace URI: ${xmlPart.namespaceUri || "(없음)"}\n\n과정:\n1. context.document.customXmlParts로 커스텀 XML 파트 컬렉션 가져오기\n2. customXmlParts.add(xmlContent)로 XML 파트 추가\n3. xmlPart.load()로 속성 로드\n4. context.sync()로 동기화\n\n참고: 커스텀 XML 파트는 문서 내부에 저장되지만 UI에서는 보이지 않습니다.\n확인 방법: 문서를 .docx로 저장한 후 압축 해제하여 customXml 폴더 확인`);
        
        setXmlContent("");
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. 모든 커스텀 XML 파트 읽기
  const readCustomXmlParts = async () => {
    try {
      await Word.run(async (context) => {
        const xmlParts = context.document.customXmlParts;
        xmlParts.load("items/id,items/namespaceUri");
        await context.sync();

        if (xmlParts.items.length === 0) {
          setResult("커스텀 XML 파트가 없습니다.\n\n과정:\n1. context.document.customXmlParts로 커스텀 XML 파트 컬렉션 가져오기\n2. customXmlParts.load()로 속성 로드\n3. context.sync()로 동기화");
          return;
        }

        const xmlList = await Promise.all(xmlParts.items.map(async (xmlPart: any) => {
          xmlPart.load("id,namespaceUri");
          const xml = xmlPart.getXml();
          await context.sync();
          return `ID: ${xmlPart.id}\nNamespace: ${xmlPart.namespaceUri || "(없음)"}\nXML: ${xml.value.substring(0, 200)}${xml.value.length > 200 ? "..." : ""}`;
        }));

        await context.sync();

        setResult(`커스텀 XML 파트 목록 (${xmlParts.items.length}개):\n\n${xmlList.join("\n\n")}\n\n과정:\n1. context.document.customXmlParts로 커스텀 XML 파트 컬렉션 가져오기\n2. customXmlParts.load()로 속성 로드\n3. 각 XML 파트의 getXml()로 XML 내용 가져오기\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 7. Add-in Settings 추가/설정
  const setAddInSetting = async () => {
    if (!settingKey.trim() || !settingValue.trim()) {
      setResult("설정 키와 값을 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const settings = context.document.settings;
        settings.add(settingKey, settingValue);
        await context.sync();

        setResult(`Add-in 설정 추가 완료!\n키: ${settingKey}\n값: ${settingValue}\n\n과정:\n1. context.document.settings로 Add-in 설정 컬렉션 가져오기\n2. settings.add(key, value)로 설정 추가\n3. context.sync()로 동기화\n\n참고: Add-in Settings는 이 Add-in 전용으로 저장되며, 다른 Add-in에서는 접근할 수 없습니다.`);
        
        setSettingKey("");
        setSettingValue("");
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 모든 Add-in Settings 읽기
  const readAddInSettings = async () => {
    try {
      await Word.run(async (context) => {
        const settings = context.document.settings;
        settings.load("items/key,items/value");
        await context.sync();

        if (settings.items.length === 0) {
          setResult("Add-in 설정이 없습니다.\n\n과정:\n1. context.document.settings로 Add-in 설정 컬렉션 가져오기\n2. settings.load()로 설정 로드\n3. context.sync()로 동기화");
          return;
        }

        const settingList = settings.items.map((setting: any, idx: number) => {
          return `${idx + 1}. ${setting.key}: ${setting.value}`;
        }).join("\n");

        setResult(`Add-in 설정 목록 (${settings.items.length}개):\n\n${settingList}\n\n과정:\n1. context.document.settings로 Add-in 설정 컬렉션 가져오기\n2. settings.load()로 설정 로드\n3. context.sync()로 동기화\n4. items 배열을 순회하여 정보 표시`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word 문서 메타데이터 및 속성</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#856404", fontSize: "14px" }}>📋 메타데이터 및 속성 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#856404" }}>
            메타데이터는 문서 내용이 아닌, 문서를 설명하거나 관리하는 정보입니다.
          </p>
          <p style={{ margin: "8px 0", color: "#856404", fontWeight: "bold" }}>📌 주요 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#856404" }}>
            <li><strong>기본 속성:</strong> 제목, 작성자, 주제 등 Word가 제공하는 기본 속성</li>
            <li><strong>커스텀 속성:</strong> 사용자가 정의한 고유한 속성 (예: 프로젝트 ID, 버전 번호)</li>
            <li><strong>커스텀 XML:</strong> 구조화된 데이터를 XML 형식으로 문서에 저장</li>
            <li><strong>Add-in Settings:</strong> 이 Add-in 전용 설정 데이터</li>
          </ul>
          <p style={{ margin: "8px 0 0 0", color: "#d32f2f", fontSize: "12px", fontStyle: "italic" }}>
            ⚠️ 확인 방법: 메타데이터는 문서 내용에는 보이지 않지만, Word UI에서 &quot;파일 &gt; 정보 &gt; 속성&quot;에서 확인할 수 있습니다.
          </p>
        </div>

        {/* 기본 속성 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #2196f3" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>📄 기본 문서 속성</h4>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={readBuiltInProperties}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              기본 속성 읽기
            </button>
            <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginTop: "10px" }}>
              <input
                type="text"
                id="titleInput"
                placeholder="제목 입력"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minWidth: "150px" }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    const input = e.target as HTMLInputElement;
                    if (input.value.trim()) {
                      setBuiltInProperty("title", input.value.trim());
                      input.value = "";
                    }
                  }
                }}
              />
              <button
                onClick={() => {
                  const input = document.getElementById("titleInput") as HTMLInputElement;
                  if (input && input.value.trim()) {
                    setBuiltInProperty("title", input.value.trim());
                    input.value = "";
                  } else {
                    setResult("제목을 입력해주세요.");
                  }
                }}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#2196f3",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                제목 설정
              </button>
              <input
                type="text"
                id="authorInput"
                placeholder="작성자 입력"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", minWidth: "150px" }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    const input = e.target as HTMLInputElement;
                    if (input.value.trim()) {
                      setBuiltInProperty("author", input.value.trim());
                      input.value = "";
                    }
                  }
                }}
              />
              <button
                onClick={() => {
                  const input = document.getElementById("authorInput") as HTMLInputElement;
                  if (input && input.value.trim()) {
                    setBuiltInProperty("author", input.value.trim());
                    input.value = "";
                  } else {
                    setResult("작성자를 입력해주세요.");
                  }
                }}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#2196f3",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                작성자 설정
              </button>
            </div>
          </div>
        </div>

        {/* 커스텀 속성 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>🔧 커스텀 속성</h4>
          <div style={{ marginBottom: "10px" }}>
            <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
              <input
                type="text"
                value={customPropName}
                onChange={(e) => setCustomPropName(e.target.value)}
                placeholder="속성 이름 (예: ProjectID)"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "150px" }}
              />
              <input
                type="text"
                value={customPropValue}
                onChange={(e) => setCustomPropValue(e.target.value)}
                placeholder="속성 값"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "150px" }}
              />
              <button
                onClick={addCustomProperty}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#4caf50",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                추가
              </button>
            </div>
            <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
              <button
                onClick={readCustomProperties}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#ff9800",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                커스텀 속성 목록
              </button>
            </div>
          </div>
        </div>

        {/* 커스텀 XML 섹션 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>📦 커스텀 XML 파트</h4>
          <div style={{ marginBottom: "10px" }}>
            <textarea
              value={xmlContent}
              onChange={(e) => setXmlContent(e.target.value)}
              placeholder='XML 내용 (예: <Meta xmlns="http://example.com/meta"><Reviewer>홍길동</Reviewer></Meta>)'
              style={{ 
                width: "100%", 
                padding: "8px", 
                border: "1px solid #ddd", 
                borderRadius: "5px", 
                minHeight: "60px",
                fontFamily: "monospace",
                fontSize: "12px"
              }}
            />
            <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginTop: "10px" }}>
              <button
                onClick={addCustomXmlPart}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#9c27b0",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                XML 파트 추가
              </button>
              <button
                onClick={readCustomXmlParts}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#ff9800",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                XML 파트 목록
              </button>
            </div>
          </div>
        </div>

        {/* Add-in Settings 섹션 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff5722" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff5722" }}>⚙️ Add-in Settings</h4>
          <div style={{ marginBottom: "10px" }}>
            <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "10px" }}>
              <input
                type="text"
                value={settingKey}
                onChange={(e) => setSettingKey(e.target.value)}
                placeholder="설정 키"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "150px" }}
              />
              <input
                type="text"
                value={settingValue}
                onChange={(e) => setSettingValue(e.target.value)}
                placeholder="설정 값"
                style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", flex: "1", minWidth: "150px" }}
              />
              <button
                onClick={setAddInSetting}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#ff5722",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                설정 추가
              </button>
            </div>
            <button
              onClick={readAddInSettings}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              설정 목록
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
          {result || "위 버튼을 클릭하여 메타데이터 및 속성 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Metadata;
