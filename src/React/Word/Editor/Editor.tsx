import React, { useState } from "react";

/* global Word */

const Editor: React.FC = () => {
  const [searchText, setSearchText] = useState("");
  const [replaceTextValue, setReplaceTextValue] = useState("");
  const [result, setResult] = useState("");

  // 1. 사용자가 선택한 텍스트 가져오기
  const selectUserSelection = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        
        if (selection.text.trim() === "") {
          setResult("사용자가 선택한 텍스트가 없습니다. Word 문서에서 텍스트를 선택해주세요.");
          return;
        }
        
        setResult(`선택된 텍스트: "${selection.text}"\n\n과정:\n1. context.document.getSelection()으로 사용자 선택 가져오기\n2. selection.load("text")로 텍스트 속성 로드\n3. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 문서 전체 선택
  const selectDocumentBody = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        
        // 문서 전체를 실제로 선택 (드래그된 것처럼 표시)
        body.select();
        await context.sync();
        
        setResult(`문서 전체 텍스트 (${body.text.length}자):\n${body.text.substring(0, 200)}${body.text.length > 200 ? "..." : ""}\n\n과정:\n1. context.document.body로 문서 본문 가져오기\n2. body.load("text")로 텍스트 속성 로드\n3. body.select()로 문서 전체 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. 텍스트 검색으로 선택
  const selectBySearch = async () => {
    if (!searchText.trim()) {
      setResult("검색할 텍스트를 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        const searchResults = context.document.body.search(searchText, {
          matchCase: false,
          matchWholeWord: false,
        });
        searchResults.load("text");
        await context.sync();
        
        if (searchResults.items.length === 0) {
          setResult(`"${searchText}"를 찾을 수 없습니다.\n\n과정:\n1. context.document.body.search("${searchText}")로 검색\n2. searchResults.load("text")로 텍스트 속성 로드\n3. context.sync()로 동기화`);
          return;
        }
        
        // 첫 번째 검색 결과를 실제로 선택 (드래그된 것처럼 표시)
        searchResults.items[0].select();
        await context.sync();
        
        const foundTexts = searchResults.items.map((item, idx) => `${idx + 1}. "${item.text}"`).join("\n");
        setResult(`찾은 텍스트 (${searchResults.items.length}개):\n${foundTexts}\n\n과정:\n1. context.document.body.search("${searchText}")로 검색\n2. searchResults.load("text")로 텍스트 속성 로드\n3. searchResults.items[0].select()로 첫 번째 결과 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 4. 첫 번째 단락 선택
  const selectFirstParagraph = async () => {
    try {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        await context.sync();
        
        if (paragraphs.items.length === 0) {
          setResult("단락이 없습니다.");
          return;
        }
        
        // 첫 번째 단락을 실제로 선택 (드래그된 것처럼 표시)
        const firstParagraph = paragraphs.items[0];
        firstParagraph.select();
        await context.sync();
        
        setResult(`첫 번째 단락: "${firstParagraph.text}"\n\n과정:\n1. context.document.body.paragraphs로 모든 단락 가져오기\n2. paragraphs.load("text")로 텍스트 속성 로드\n3. paragraph.select()로 단락 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. 특정 범위(Range)로 선택
  const selectByRange = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("text");
        await context.sync();
        
        if (paragraphs.items.length === 0) {
          setResult("단락이 없습니다.");
          return;
        }
        
        // 첫 번째 단락의 Range 가져오기
        const firstParagraph = paragraphs.items[0];
        const simpleRange = firstParagraph.getRange();
        simpleRange.load("text");
        await context.sync();
        
        // Range를 실제로 선택 (드래그된 것처럼 표시)
        simpleRange.select();
        await context.sync();
        
        setResult(`Range로 선택한 텍스트: "${simpleRange.text.substring(0, 50)}${simpleRange.text.length > 50 ? "..." : ""}"\n\n과정:\n1. paragraph.getRange()로 Range 객체 생성\n2. range.load("text")로 텍스트 속성 로드\n3. range.select()로 Range 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. 섹션(Section)으로 선택
  const selectBySection = async () => {
    try {
      await Word.run(async (context) => {
        const sections = context.document.sections;
        sections.load("body");
        await context.sync();
        
        if (sections.items.length === 0) {
          setResult("섹션이 없습니다.");
          return;
        }
        
        const firstSection = sections.items[0];
        firstSection.body.load("text");
        await context.sync();
        
        // 섹션 본문을 실제로 선택 (드래그된 것처럼 표시)
        firstSection.body.select();
        await context.sync();
        
        setResult(`첫 번째 섹션 텍스트: "${firstSection.body.text.substring(0, 100)}${firstSection.body.text.length > 100 ? "..." : ""}"\n\n과정:\n1. context.document.sections로 모든 섹션 가져오기\n2. section.body.load("text")로 본문 텍스트 로드\n3. section.body.select()로 섹션 본문 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 7. 북마크(Bookmark)로 선택
  const selectByBookmark = async () => {
    try {
      await Word.run(async (context) => {
        const bookmarks = context.document.bookmarks;
        bookmarks.load("name");
        await context.sync();
        
        if (bookmarks.items.length === 0) {
          setResult("북마크가 없습니다. Word 문서에 북마크를 추가해주세요.\n\n과정:\n1. context.document.bookmarks로 모든 북마크 가져오기\n2. bookmarks.load('name')로 이름 속성 로드\n3. context.sync()로 동기화\n\n참고: 북마크의 텍스트는 bookmark.name으로 이름만 가져올 수 있으며, 실제 텍스트는 다른 방법으로 접근해야 합니다.");
          return;
        }
        
        const firstBookmark = bookmarks.items[0];
        firstBookmark.load("range");
        await context.sync();
        
        const bookmarkRange = firstBookmark.range;
        bookmarkRange.load("text");
        await context.sync();
        
        // 북마크 범위를 실제로 선택 (드래그된 것처럼 표시)
        bookmarkRange.select();
        await context.sync();
        
        setResult(`북마크 "${firstBookmark.name}"의 텍스트: "${bookmarkRange.text}"\n\n과정:\n1. context.document.bookmarks로 모든 북마크 가져오기\n2. bookmark.load("range")로 Range 객체 로드\n3. bookmark.range.select()로 북마크 범위 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 콘텐츠 컨트롤(Content Control)로 선택
  const selectByContentControl = async () => {
    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("title");
        await context.sync();
        
        if (contentControls.items.length === 0) {
          setResult("콘텐츠 컨트롤이 없습니다. Word 문서에 콘텐츠 컨트롤을 추가해주세요.\n\n과정:\n1. context.document.contentControls로 모든 콘텐츠 컨트롤 가져오기\n2. contentControls.load('title')로 제목 속성 로드\n3. context.sync()로 동기화");
          return;
        }
        
        const firstControl = contentControls.items[0];
        firstControl.load("text,title");
        await context.sync();
        
        // 콘텐츠 컨트롤을 실제로 선택 (드래그된 것처럼 표시)
        const controlRange = firstControl.getRange();
        controlRange.select();
        await context.sync();
        
        const controlTitle = firstControl.title || "(제목 없음)";
        setResult(`콘텐츠 컨트롤 "${controlTitle}"의 텍스트: "${firstControl.text}"\n\n과정:\n1. context.document.contentControls로 모든 콘텐츠 컨트롤 가져오기\n2. contentControl.load("text,title")로 텍스트와 제목 속성 로드\n3. contentControl.getRange().select()로 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 9. 특정 인덱스의 단락 선택
  const selectParagraphByIndex = async () => {
    try {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        await context.sync();
        
        if (paragraphs.items.length === 0) {
          setResult("단락이 없습니다.");
          return;
        }
        
        // 마지막 단락 선택
        const lastIndex = paragraphs.items.length - 1;
        const lastParagraph = paragraphs.items[lastIndex];
        lastParagraph.load("text");
        await context.sync();
        
        // 마지막 단락을 실제로 선택 (드래그된 것처럼 표시)
        lastParagraph.select();
        await context.sync();
        
        setResult(`마지막 단락 (인덱스 ${lastIndex}): "${lastParagraph.text}"\n\n과정:\n1. context.document.body.paragraphs로 모든 단락 가져오기\n2. paragraphs.items[index]로 특정 인덱스의 단락 접근\n3. paragraph.select()로 단락 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 10. 문장(Sentence)으로 선택 - 첫 번째 단락의 텍스트를 문장 단위로 분리
  const selectBySentence = async () => {
    try {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        await context.sync();
        
        if (paragraphs.items.length === 0) {
          setResult("단락이 없습니다.");
          return;
        }
        
        const firstParagraph = paragraphs.items[0];
        const paragraphText = firstParagraph.text;
        const sentences = paragraphText.split(/[.!?]\s+/).filter(s => s.trim());
        
        if (sentences.length === 0) {
          setResult("문장이 없습니다.");
          return;
        }
        
        // 첫 번째 문장을 검색하여 선택
        const firstSentenceText = sentences[0];
        const searchResults = firstParagraph.search(firstSentenceText, {
          matchCase: false,
          matchWholeWord: false,
        });
        searchResults.load("text");
        await context.sync();
        
        if (searchResults.items.length > 0) {
          searchResults.items[0].select();
          await context.sync();
        }
        
        setResult(`첫 번째 문장: "${firstSentenceText}"\n\n과정:\n1. paragraph.text로 텍스트 가져오기\n2. 텍스트를 문장 단위로 분리 (정규식 사용)\n3. paragraph.search()로 문장 검색 후 select()로 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화\n\n참고: Word API에는 직접적인 sentences 속성이 없어 텍스트를 파싱하여 사용`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 11. 단어(Word)로 선택 - 첫 번째 단락의 텍스트를 단어 단위로 분리
  const selectByWord = async () => {
    try {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("text");
        await context.sync();
        
        if (paragraphs.items.length === 0) {
          setResult("단락이 없습니다.");
          return;
        }
        
        const firstParagraph = paragraphs.items[0];
        const paragraphText = firstParagraph.text;
        const words = paragraphText.split(/\s+/).filter(w => w.trim());
        
        if (words.length === 0) {
          setResult("단어가 없습니다.");
          return;
        }
        
        // 첫 번째 단어를 검색하여 선택
        const firstWordText = words[0].replace(/[.,!?;:]/g, "");
        const searchResults = firstParagraph.search(firstWordText, {
          matchCase: false,
          matchWholeWord: true,
        });
        searchResults.load("text");
        await context.sync();
        
        if (searchResults.items.length > 0) {
          searchResults.items[0].select();
          await context.sync();
        }
        
        setResult(`첫 번째 단어: "${words[0]}"\n\n과정:\n1. paragraph.text로 텍스트 가져오기\n2. 텍스트를 단어 단위로 분리 (공백 기준)\n3. paragraph.search()로 단어 검색 후 select()로 선택 (드래그된 것처럼 표시)\n4. context.sync()로 동기화\n\n참고: Word API에는 직접적인 words 속성이 없어 텍스트를 파싱하여 사용`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 12. 표(Table) 내부 텍스트 선택
  const selectByTable = async () => {
    try {
      await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("rowCount,columnCount");
        await context.sync();
        
        if (tables.items.length === 0) {
          setResult("표가 없습니다. Word 문서에 표를 추가해주세요.\n\n과정:\n1. context.document.body.tables로 모든 표 가져오기\n2. tables.load('rowCount,columnCount')로 속성 로드\n3. context.sync()로 동기화");
          return;
        }
        
        const firstTable = tables.items[0];
        const firstRow = firstTable.rows.getFirst();
        firstRow.load("cells");
        await context.sync();
        
        const cells = firstRow.cells;
        cells.load("body");
        await context.sync();
        
        const cellTexts = cells.items.map(cell => {
          cell.body.load("text");
          return cell.body.text;
        });
        await context.sync();
        
        // 첫 번째 셀의 본문을 실제로 선택 (드래그된 것처럼 표시)
        const firstCell = cells.items[0];
        firstCell.body.select();
        await context.sync();
        
        const tableText = cellTexts.join(" | ");
        setResult(`첫 번째 표의 첫 번째 행: ${tableText}\n\n과정:\n1. context.document.body.tables로 모든 표 가져오기\n2. table.rows.getFirst()로 첫 번째 행 가져오기\n3. row.cells로 셀들 가져오기\n4. cell.body.load("text")로 텍스트 로드\n5. cell.body.select()로 첫 번째 셀 선택 (드래그된 것처럼 표시)\n6. context.sync()로 동기화`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 텍스트 대치 기능
  const replaceSelectedText = async (selectionMethod: string) => {
    if (!replaceTextValue.trim()) {
      setResult("대치할 텍스트를 입력해주세요.");
      return;
    }

    try {
      await Word.run(async (context) => {
        let targetRange: Word.Range;

        if (selectionMethod === "user") {
          // 사용자 선택 텍스트 대치
          const selection = context.document.getSelection();
          selection.load("text");
          await context.sync();
          
          if (selection.text.trim() === "") {
            setResult("사용자가 선택한 텍스트가 없습니다.");
            return;
          }
          
          targetRange = selection;
        } else if (selectionMethod === "search") {
          // 검색된 모든 텍스트 대치
          if (!searchText.trim()) {
            setResult("검색할 텍스트를 입력해주세요.");
            return;
          }
          
          const searchResults = context.document.body.search(searchText, {
            matchCase: false,
            matchWholeWord: false,
          });
          searchResults.load("text");
          await context.sync();
          
          if (searchResults.items.length === 0) {
            setResult(`"${searchText}"를 찾을 수 없습니다.`);
            return;
          }
          
          // 역순으로 처리하여 인덱스 변경 문제 방지
          const replacedCount = searchResults.items.length;
          for (let i = searchResults.items.length - 1; i >= 0; i--) {
            searchResults.items[i].insertText(replaceTextValue, Word.InsertLocation.replace);
          }
          await context.sync();
          
          setResult(`텍스트 대치 완료! (${replacedCount}개)\n"${searchText}" → "${replaceTextValue}"`);
          return;
        } else {
          // 문서 전체는 대치하지 않음 (너무 위험)
          setResult("문서 전체 대치는 위험하므로 지원하지 않습니다.");
          return;
        }

        const originalText = targetRange.text;
        targetRange.insertText(replaceTextValue, Word.InsertLocation.replace);
        await context.sync();
        
        setResult(`텍스트 대치 완료!\n"${originalText}" → "${replaceTextValue}"`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word 텍스트 선택 방법 테스트</h3>
        <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", marginBottom: "15px" }}>
          <button
            onClick={selectUserSelection}
            style={{
              padding: "6px 12px",
              backgroundColor: "#2196f3",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            1. 사용자 선택
          </button>
          <button
            onClick={selectDocumentBody}
            style={{
              padding: "6px 12px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            2. 문서 전체
          </button>
          <button
            onClick={selectFirstParagraph}
            style={{
              padding: "6px 12px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            3. 첫 단락
          </button>
          <button
            onClick={selectParagraphByIndex}
            style={{
              padding: "6px 12px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            4. 마지막 단락
          </button>
          <button
            onClick={selectByRange}
            style={{
              padding: "6px 12px",
              backgroundColor: "#9c27b0",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            5. Range
          </button>
          <button
            onClick={selectBySection}
            style={{
              padding: "6px 12px",
              backgroundColor: "#00bcd4",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            6. 섹션
          </button>
          <button
            onClick={selectByBookmark}
            style={{
              padding: "6px 12px",
              backgroundColor: "#795548",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            7. 북마크
          </button>
          <button
            onClick={selectByContentControl}
            style={{
              padding: "6px 12px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            8. 콘텐츠 컨트롤
          </button>
          <button
            onClick={selectBySentence}
            style={{
              padding: "6px 12px",
              backgroundColor: "#e91e63",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            9. 문장
          </button>
          <button
            onClick={selectByWord}
            style={{
              padding: "6px 12px",
              backgroundColor: "#e91e63",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            10. 단어
          </button>
          <button
            onClick={selectByTable}
            style={{
              padding: "6px 12px",
              backgroundColor: "#3f51b5",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            11. 표
          </button>
        </div>
        <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
          <input
            type="text"
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
            placeholder="검색할 텍스트"
            style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "200px" }}
          />
          <button
            onClick={selectBySearch}
            style={{
              padding: "6px 12px",
              backgroundColor: "#f44336",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
              fontSize: "11px",
            }}
          >
            12. 검색
          </button>
        </div>
      </div>

      <div style={{ flex: 1, overflowY: "auto", padding: "15px" }}>
        <h4 style={{ marginTop: 0 }}>결과 및 과정 설명</h4>
        <pre style={{
          backgroundColor: "#f5f5f5",
          padding: "15px",
          borderRadius: "5px",
          whiteSpace: "pre-wrap",
          fontFamily: "monospace",
          fontSize: "12px",
          lineHeight: "1.5",
        }}>
          {result || "위 버튼을 클릭하여 다양한 텍스트 선택 방법을 테스트해보세요."}
        </pre>
      </div>

      <div style={{ padding: "15px", borderTop: "1px solid #ddd", backgroundColor: "#fff3cd" }}>
        <h4 style={{ margin: "0 0 10px 0" }}>텍스트 대치</h4>
        <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
          <input
            type="text"
            value={replaceTextValue}
            onChange={(e) => setReplaceTextValue(e.target.value)}
            placeholder="대치할 텍스트"
            style={{ padding: "8px", border: "1px solid #ddd", borderRadius: "5px", width: "200px" }}
          />
          <button
            onClick={() => replaceSelectedText("user")}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            사용자 선택 대치
          </button>
          <button
            onClick={() => replaceSelectedText("search")}
            style={{
              padding: "8px 16px",
              backgroundColor: "#607d8b",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            검색 결과 대치
          </button>
        </div>
      </div>
    </div>
  );
};

export default Editor;
