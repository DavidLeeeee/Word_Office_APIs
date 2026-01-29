import React, { Component, ErrorInfo, ReactNode } from "react";
import { addinLogger } from "./AddinLogger";

interface Props {
  children: ReactNode;
  fallback?: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
  errorInfo: ErrorInfo | null;
}

/**
 * React Error Boundary 컴포넌트
 * Add-in의 React 렌더링 단계에서 발생하는 에러를 캡처
 */
export class ErrorBoundary extends Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = {
      hasError: false,
      error: null,
      errorInfo: null,
    };
  }

  static getDerivedStateFromError(error: Error): State {
    return {
      hasError: true,
      error,
      errorInfo: null,
    };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    // 에러 로깅
    addinLogger.error("React 컴포넌트 렌더링 에러", error, "ErrorBoundary");

    this.setState({
      error,
      errorInfo,
    });
  }

  render() {
    if (this.state.hasError) {
      // 커스텀 fallback UI가 있으면 사용
      if (this.props.fallback) {
        return this.props.fallback;
      }

      // 기본 에러 UI
      return (
        <div
          style={{
            padding: "20px",
            backgroundColor: "#fff3cd",
            border: "2px solid #ffc107",
            borderRadius: "5px",
            margin: "20px",
          }}
        >
          <h2 style={{ color: "#856404", marginTop: 0 }}>⚠️ Add-in 로드 오류</h2>
          <p style={{ color: "#856404", marginBottom: "15px" }}>
            Add-in을 로드하는 중 오류가 발생했습니다.
          </p>
          {this.state.error && (
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#856404" }}>에러 메시지:</strong>
              <pre
                style={{
                  backgroundColor: "#fff",
                  padding: "10px",
                  borderRadius: "4px",
                  overflow: "auto",
                  fontSize: "12px",
                  color: "#d32f2f",
                }}
              >
                {this.state.error.toString()}
              </pre>
            </div>
          )}
          {this.state.errorInfo && (
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#856404" }}>컴포넌트 스택:</strong>
              <pre
                style={{
                  backgroundColor: "#fff",
                  padding: "10px",
                  borderRadius: "4px",
                  overflow: "auto",
                  fontSize: "11px",
                  maxHeight: "200px",
                }}
              >
                {this.state.errorInfo.componentStack}
              </pre>
            </div>
          )}
          <div style={{ marginTop: "20px" }}>
            <button
              onClick={() => {
                addinLogger.downloadLogs();
              }}
              style={{
                padding: "10px 20px",
                backgroundColor: "#ffc107",
                color: "#000",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                marginRight: "10px",
                fontWeight: "bold",
              }}
            >
              로그 다운로드
            </button>
            <button
              onClick={() => {
                window.location.reload();
              }}
              style={{
                padding: "10px 20px",
                backgroundColor: "#1976d2",
                color: "#fff",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                fontWeight: "bold",
              }}
            >
              페이지 새로고침
            </button>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
