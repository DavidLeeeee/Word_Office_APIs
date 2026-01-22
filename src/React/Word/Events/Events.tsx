import React, { useState, useEffect } from "react";

/* global Office, Word */

const Events: React.FC = () => {
  const [result, setResult] = useState("");
  const [eventLog, setEventLog] = useState<string[]>([]);
  const [isListening, setIsListening] = useState(false);
  const [isBindingListening, setIsBindingListening] = useState(false);
  const [bindingId, setBindingId] = useState<string | null>(null);
  const [bindingData, setBindingData] = useState<string>("");
  const [bindingTextToSet, setBindingTextToSet] = useState<string>("");

  // 이벤트 로그에 추가
  const addEventLog = (message: string) => {
    const timestamp = new Date().toLocaleTimeString();
    const logEntry = `[${timestamp}] ${message}`;
    setEventLog((prev) => [logEntry, ...prev].slice(0, 50)); // 최대 50개만 유지
    setResult(logEntry);
  };

  // 1. Selection Changed 이벤트 등록
  const startSelectionChangedListener = () => {
    try {
      if (!Office || !Office.context || !Office.context.document) {
        addEventLog("오류: Office.context.document을 사용할 수 없습니다.");
        return;
      }

      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        (eventArgs: Office.DocumentSelectionChangedEventArgs) => {
          addEventLog("✅ 선택 변경 이벤트 감지됨!");
          
          // 선택된 텍스트 가져오기
          Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            addEventLog(`   선택된 텍스트: "${selection.text.substring(0, 50)}${selection.text.length > 50 ? "..." : ""}"`);
          }).catch((error) => {
            addEventLog(`   오류: ${error.message}`);
          });
        },
        (result: Office.AsyncResult<void>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            setIsListening(true);
            addEventLog("✅ Selection Changed 이벤트 리스너 등록 완료!");
            addEventLog("   이제 Word 문서에서 텍스트를 선택하면 이벤트가 감지됩니다.");
          } else {
            addEventLog(`❌ 이벤트 등록 실패: ${result.error?.message || "알 수 없는 오류"}`);
          }
        }
      );
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. Binding 생성 및 이벤트 등록
  const createBindingAndListen = () => {
    try {
      if (!Office || !Office.context || !Office.context.document) {
        addEventLog("오류: Office.context.document을 사용할 수 없습니다.");
        return;
      }

      // 현재 선택된 영역을 Binding으로 생성
      Office.context.document.bindings.addFromSelectionAsync(
        Office.BindingType.Text,
        { id: `EventTestBinding_${Date.now()}` },
        (result: Office.AsyncResult<Office.Binding>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            const binding = result.value;
            setBindingId(binding.id);
            addEventLog(`✅ Binding 생성 완료! ID: ${binding.id}`);
            addEventLog("   이제 Binding 내부의 텍스트를 수정하거나 선택해보세요.");

            // BindingDataChanged 이벤트 등록
            binding.addHandlerAsync(
              Office.EventType.BindingDataChanged,
              (eventArgs: Office.BindingDataChangedEventArgs) => {
                addEventLog("✅ Binding Data Changed 이벤트 감지됨!");
                
                // 변경된 데이터 읽기
                binding.getDataAsync({ coercionType: Office.CoercionType.Text }, (dataResult: Office.AsyncResult<string>) => {
                  if (dataResult.status === Office.AsyncResultStatus.Succeeded) {
                    addEventLog(`   변경된 데이터: "${dataResult.value?.substring(0, 50)}${(dataResult.value?.length || 0) > 50 ? "..." : ""}"`);
                  }
                });
              },
              (handlerResult: Office.AsyncResult<void>) => {
                if (handlerResult.status === Office.AsyncResultStatus.Succeeded) {
                  addEventLog("✅ BindingDataChanged 이벤트 리스너 등록 완료!");
                } else {
                  addEventLog(`❌ BindingDataChanged 등록 실패: ${handlerResult.error?.message}`);
                }
              }
            );

            // BindingSelectionChanged 이벤트 등록
            binding.addHandlerAsync(
              Office.EventType.BindingSelectionChanged,
              (eventArgs: Office.BindingSelectionChangedEventArgs) => {
                addEventLog("✅ Binding Selection Changed 이벤트 감지됨!");
              },
              (handlerResult: Office.AsyncResult<void>) => {
                if (handlerResult.status === Office.AsyncResultStatus.Succeeded) {
                  addEventLog("✅ BindingSelectionChanged 이벤트 리스너 등록 완료!");
                  setIsBindingListening(true);
                } else {
                  addEventLog(`❌ BindingSelectionChanged 등록 실패: ${handlerResult.error?.message}`);
                }
              }
            );
          } else {
            addEventLog(`❌ Binding 생성 실패: ${result.error?.message || "알 수 없는 오류"}`);
            addEventLog("   참고: Word 문서에서 텍스트를 선택한 후 다시 시도해주세요.");
          }
        }
      );
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 3. Binding 이벤트 리스너 제거
  const removeBindingListeners = () => {
    try {
      if (!bindingId || !Office || !Office.context || !Office.context.document) {
        addEventLog("⚠️ 제거할 Binding이 없습니다.");
        return;
      }

      Office.context.document.bindings.getByIdAsync(
        bindingId,
        (result: Office.AsyncResult<Office.Binding>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            const binding = result.value;
            
            // BindingDataChanged 리스너 제거
            binding.removeHandlerAsync(
              Office.EventType.BindingDataChanged,
              { handler: () => {} },
              (removeResult: Office.AsyncResult<void>) => {
                // BindingSelectionChanged 리스너 제거
                binding.removeHandlerAsync(
                  Office.EventType.BindingSelectionChanged,
                  { handler: () => {} },
                  (removeResult2: Office.AsyncResult<void>) => {
                    // Binding 자체도 제거
                    Office.context.document.bindings.releaseByIdAsync(
                      bindingId,
                      (releaseResult: Office.AsyncResult<void>) => {
                        setIsBindingListening(false);
                        setBindingId(null);
                        addEventLog("✅ Binding 이벤트 리스너 및 Binding 제거 완료!");
                      }
                    );
                  }
                );
              }
            );
          } else {
            addEventLog(`❌ Binding을 찾을 수 없습니다: ${result.error?.message || "알 수 없는 오류"}`);
            setIsBindingListening(false);
            setBindingId(null);
          }
        }
      );
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
      setIsBindingListening(false);
      setBindingId(null);
    }
  };

  // 4. 이벤트 리스너 제거
  const removeAllListeners = () => {
    try {
      if (!Office || !Office.context || !Office.context.document) {
        addEventLog("오류: Office.context.document을 사용할 수 없습니다.");
        return;
      }

      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: () => {} },
        (result: Office.AsyncResult<void>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            setIsListening(false);
            addEventLog("✅ 이벤트 리스너 제거 완료!");
          } else {
            addEventLog(`❌ 리스너 제거 실패: ${result.error?.message || "알 수 없는 오류"}`);
          }
        }
      );
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 5. Binding 데이터 읽기
  const readBindingData = () => {
    if (!bindingId) {
      addEventLog("⚠️ 먼저 Binding을 생성해주세요.");
      return;
    }

    try {
      Office.context.document.bindings.getByIdAsync(
        bindingId,
        (result: Office.AsyncResult<Office.Binding>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            const binding = result.value;
            
            // Text 형식으로 읽기
            binding.getDataAsync({ coercionType: Office.CoercionType.Text }, (dataResult: Office.AsyncResult<string>) => {
              if (dataResult.status === Office.AsyncResultStatus.Succeeded) {
                setBindingData(dataResult.value || "");
                addEventLog(`✅ Binding 데이터 읽기 완료!\n데이터: "${dataResult.value?.substring(0, 100)}${(dataResult.value?.length || 0) > 100 ? "..." : ""}"`);
              } else {
                addEventLog(`❌ 데이터 읽기 실패: ${dataResult.error?.message || "알 수 없는 오류"}`);
              }
            });
          } else {
            addEventLog(`❌ Binding을 찾을 수 없습니다: ${result.error?.message || "알 수 없는 오류"}`);
          }
        }
      );
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 6. Binding 데이터 쓰기
  const writeBindingData = () => {
    if (!bindingId) {
      addEventLog("⚠️ 먼저 Binding을 생성해주세요.");
      return;
    }

    if (!bindingTextToSet.trim()) {
      addEventLog("⚠️ 입력할 텍스트를 입력해주세요.");
      return;
    }

    try {
      Office.context.document.bindings.getByIdAsync(
        bindingId,
        (result: Office.AsyncResult<Office.Binding>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            const binding = result.value;
            
            binding.setDataAsync(
              bindingTextToSet,
              { coercionType: Office.CoercionType.Text },
              (setResult: Office.AsyncResult<void>) => {
                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                  addEventLog(`✅ Binding 데이터 쓰기 완료!\n작성한 데이터: "${bindingTextToSet}"`);
                  setBindingTextToSet("");
                  // 자동으로 다시 읽기
                  setTimeout(() => readBindingData(), 500);
                } else {
                  addEventLog(`❌ 데이터 쓰기 실패: ${setResult.error?.message || "알 수 없는 오류"}`);
                }
              }
            );
          } else {
            addEventLog(`❌ Binding을 찾을 수 없습니다: ${result.error?.message || "알 수 없는 오류"}`);
          }
        }
      );
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 7. Binding 정보 확인
  const getBindingInfo = () => {
    if (!bindingId) {
      addEventLog("⚠️ 먼저 Binding을 생성해주세요.");
      return;
    }

    try {
      Office.context.document.bindings.getByIdAsync(
        bindingId,
        (result: Office.AsyncResult<Office.Binding>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            const binding = result.value;
            const info = `📋 Binding 정보:\n\nID: ${binding.id}\nType: ${binding.type}\nDocument: ${binding.document ? "연결됨" : "없음"}\n\n과정:\n1. Office.context.document.bindings.getByIdAsync()로 Binding 가져오기\n2. binding.id, type, document 속성 확인`;
            addEventLog(info);
          } else {
            addEventLog(`❌ Binding을 찾을 수 없습니다: ${result.error?.message || "알 수 없는 오류"}`);
          }
        }
      );
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 8. 지원되는 이벤트 타입 확인
  const checkSupportedEvents = () => {
    try {
      if (!Office || !Office.EventType) {
        addEventLog("오류: Office.EventType을 사용할 수 없습니다.");
        return;
      }

      const eventTypes = [
        "DocumentSelectionChanged",
        "DocumentActiveViewChanged",
        "BindingDataChanged",
        "BindingSelectionChanged",
      ];

      let supportedEvents = "📋 지원 가능한 이벤트 타입:\n\n";
      eventTypes.forEach((eventType) => {
        const eventValue = (Office.EventType as any)[eventType];
        if (eventValue) {
          supportedEvents += `✅ ${eventType}: ${eventValue}\n`;
        } else {
          supportedEvents += `❌ ${eventType}: 지원되지 않음\n`;
        }
      });

      addEventLog(supportedEvents);
    } catch (error) {
      addEventLog(`❌ 오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 컴포넌트 언마운트 시 리스너 정리
  useEffect(() => {
    return () => {
      if (isListening) {
        removeAllListeners();
      }
      if (isBindingListening) {
        removeBindingListeners();
      }
    };
  }, [isListening, isBindingListening]);

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Word 이벤트 감지</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#856404", fontSize: "14px" }}>📝 이벤트 감지 기능 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#856404" }}>
            Word Add-in에서는 <strong>제한적인 이벤트</strong>만 감지할 수 있습니다.
          </p>
          <p style={{ margin: "8px 0", color: "#856404", fontWeight: "bold" }}>✅ 지원되는 이벤트:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#856404" }}>
            <li><strong>DocumentSelectionChanged</strong>: 사용자가 텍스트 선택을 변경할 때</li>
            <li><strong>BindingDataChanged</strong>: Binding(Content Control)의 데이터가 변경될 때</li>
            <li><strong>BindingSelectionChanged</strong>: Binding(Content Control)의 선택이 변경될 때</li>
          </ul>
          <p style={{ margin: "8px 0", color: "#856404", fontSize: "12px" }}>
            ❌ <strong>DocumentActiveViewChanged</strong>: Word에서는 지원되지 않음 (PowerPoint 전용)
          </p>
          <p style={{ margin: "8px 0", color: "#d32f2f", fontSize: "12px", fontStyle: "italic" }}>
            ⚠️ 제약사항: <strong>붙여넣기, 저장, 삭제, 입력</strong> 등의 이벤트는 Word JavaScript API에서 직접 지원되지 않습니다.
            <br />
            이러한 이벤트를 감지하려면 <strong>폴링(polling)</strong> 방식이나 다른 방법을 사용해야 합니다.
          </p>
        </div>

        {/* 이벤트 테스트 버튼들 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #2196f3" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#2196f3" }}>🔍 이벤트 테스트</h4>
          
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={checkSupportedEvents}
              style={{
                padding: "8px 16px",
                backgroundColor: "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              지원 이벤트 확인
            </button>
            <button
              onClick={startSelectionChangedListener}
              disabled={isListening}
              style={{
                padding: "8px 16px",
                backgroundColor: isListening ? "#ccc" : "#4caf50",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: isListening ? "not-allowed" : "pointer",
              }}
            >
              Selection Changed 시작
            </button>
            <button
              onClick={createBindingAndListen}
              disabled={isBindingListening}
              style={{
                padding: "8px 16px",
                backgroundColor: isBindingListening ? "#ccc" : "#ff9800",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: isBindingListening ? "not-allowed" : "pointer",
              }}
            >
              Binding 이벤트 시작
            </button>
            <button
              onClick={removeAllListeners}
              disabled={!isListening}
              style={{
                padding: "8px 16px",
                backgroundColor: !isListening ? "#ccc" : "#f44336",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: !isListening ? "not-allowed" : "pointer",
              }}
            >
              Selection 리스너 제거
            </button>
            <button
              onClick={removeBindingListeners}
              disabled={!isBindingListening}
              style={{
                padding: "8px 16px",
                backgroundColor: !isBindingListening ? "#ccc" : "#e91e63",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: !isBindingListening ? "not-allowed" : "pointer",
              }}
            >
              Binding 리스너 제거
            </button>
          </div>

          <div style={{ fontSize: "12px", color: "#666", marginTop: "5px" }}>
            {isListening && (
              <div style={{ color: "#4caf50", fontWeight: "bold", marginBottom: "5px" }}>
                ✅ Selection Changed 리스너 활성화됨 - Word 문서에서 텍스트를 선택해보세요!
              </div>
            )}
            {isBindingListening && (
              <div style={{ color: "#ff9800", fontWeight: "bold" }}>
                ✅ Binding 이벤트 리스너 활성화됨 - Content Control의 내용을 수정하거나 선택해보세요!
              </div>
            )}
          </div>
        </div>

        {/* Binding 작업 섹션 */}
        {bindingId && (
          <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
            <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 Binding 작업</h4>
            
            <div style={{ marginBottom: "10px" }}>
              <div style={{ fontSize: "12px", color: "#666", marginBottom: "5px" }}>
                Binding ID: <strong>{bindingId}</strong>
              </div>
            </div>

            <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
              <button
                onClick={getBindingInfo}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#9c27b0",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                Binding 정보 확인
              </button>
              <button
                onClick={readBindingData}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#2196f3",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: "pointer",
                }}
              >
                데이터 읽기
              </button>
            </div>

            <div style={{ marginBottom: "10px" }}>
              <div style={{ fontSize: "12px", color: "#666", marginBottom: "5px" }}>
                현재 Binding 데이터:
              </div>
              <div style={{
                backgroundColor: "#f5f5f5",
                padding: "10px",
                borderRadius: "5px",
                fontSize: "12px",
                minHeight: "40px",
                maxHeight: "100px",
                overflowY: "auto",
                border: "1px solid #ddd",
              }}>
                {bindingData || "(데이터 없음 - '데이터 읽기' 버튼을 클릭하세요)"}
              </div>
            </div>

            <div style={{ marginBottom: "10px" }}>
              <label style={{ display: "block", fontSize: "12px", color: "#666", marginBottom: "5px" }}>
                Binding에 쓸 텍스트:
              </label>
              <textarea
                value={bindingTextToSet}
                onChange={(e) => setBindingTextToSet(e.target.value)}
                placeholder="Binding에 쓸 텍스트를 입력하세요..."
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "5px",
                  fontSize: "12px",
                  minHeight: "60px",
                  resize: "vertical",
                }}
              />
              <button
                onClick={writeBindingData}
                disabled={!bindingTextToSet.trim()}
                style={{
                  marginTop: "5px",
                  padding: "8px 16px",
                  backgroundColor: bindingTextToSet.trim() ? "#4caf50" : "#ccc",
                  color: "#fff",
                  border: "none",
                  borderRadius: "5px",
                  cursor: bindingTextToSet.trim() ? "pointer" : "not-allowed",
                }}
              >
                데이터 쓰기
              </button>
            </div>

            <div style={{ fontSize: "11px", color: "#999", marginTop: "10px", padding: "10px", backgroundColor: "#f9f9f9", borderRadius: "5px" }}>
              <strong>💡 Binding 작업 설명:</strong><br />
              • <strong>데이터 읽기</strong>: Binding 영역의 현재 텍스트를 읽어옵니다.<br />
              • <strong>데이터 쓰기</strong>: Binding 영역에 새로운 텍스트를 씁니다.<br />
              • <strong>정보 확인</strong>: Binding의 ID, Type 등 정보를 확인합니다.<br />
              • Binding 내부의 텍스트를 수정하면 <strong>BindingDataChanged</strong> 이벤트가 발생합니다.
            </div>
          </div>
        )}

        {/* 현재 결과 */}
        <div style={{ padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>📊 최근 이벤트</h4>
          <div style={{
            backgroundColor: "#f5f5f5",
            padding: "10px",
            borderRadius: "5px",
            maxHeight: "150px",
            overflowY: "auto",
            fontSize: "12px",
            fontFamily: "monospace",
          }}>
            {eventLog.length === 0 ? (
              <div style={{ color: "#999" }}>이벤트 로그가 비어있습니다.</div>
            ) : (
              eventLog.map((log, index) => (
                <div key={index} style={{ marginBottom: "5px", color: "#333" }}>
                  {log}
                </div>
              ))
            )}
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
          {result || "위 버튼을 클릭하여 이벤트 감지 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Events;
