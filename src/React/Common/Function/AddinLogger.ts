/* global Office */

/**
 * Add-in 초기화 및 로드 단계 에러 로깅 유틸리티
 * Add-in 자체가 로드되지 않거나 초기화 단계에서 발생하는 에러를 캡처
 */

interface LogEntry {
  timestamp: string;
  level: "error" | "warn" | "info";
  message: string;
  error?: {
    name?: string;
    message?: string;
    code?: string;
    stack?: string;
  };
  context?: string;
  userAgent?: string;
  officeHost?: string;
  officeVersion?: string;
}

class AddinLogger {
  private logs: LogEntry[] = [];
  private maxLogs = 200;

  constructor() {
    // 초기화 정보 기록
    this.log("info", "Add-in 초기화 시작", {
      userAgent: navigator.userAgent,
      officeHost: Office?.context?.host ? String(Office.context.host) : "Unknown",
      officeVersion: Office?.context?.platform ? String(Office.context.platform) : "Unknown",
    });
  }

  /**
   * 로그 기록
   */
  log(
    level: "error" | "warn" | "info",
    message: string,
    context?: {
      error?: any;
      context?: string;
      userAgent?: string;
      officeHost?: string;
      officeVersion?: string;
    }
  ): void {
    const logEntry: LogEntry = {
      timestamp: new Date().toISOString(),
      level,
      message,
      context: context?.context,
      userAgent: context?.userAgent || navigator.userAgent,
      officeHost: context?.officeHost || (Office?.context?.host ? String(Office.context.host) : "Unknown"),
      officeVersion: context?.officeVersion || (Office?.context?.platform ? String(Office.context.platform) : "Unknown"),
    };

    if (context?.error) {
      logEntry.error = {
        name: context.error?.name,
        message: context.error?.message || String(context.error),
        code: context.error?.code,
        stack: context.error?.stack,
      };
    }

    this.logs.push(logEntry);

    // 최대 로그 수 제한
    if (this.logs.length > this.maxLogs) {
      this.logs.shift();
    }

    // 콘솔에도 출력
    const consoleMethod = level === "error" ? console.error : level === "warn" ? console.warn : console.log;
    consoleMethod(`[AddinLogger ${level.toUpperCase()}]`, message, context?.error || "");

    // 에러인 경우 즉시 서버로 전송 시도
    if (level === "error") {
      this.sendToServer(logEntry).catch((err) => {
        console.error("로그 전송 실패:", err);
      });
    }
  }

  /**
   * 에러 로그
   */
  error(message: string, error?: any, context?: string): void {
    this.log("error", message, { error, context });
  }

  /**
   * 경고 로그
   */
  warn(message: string, context?: string): void {
    this.log("warn", message, { context });
  }

  /**
   * 정보 로그
   */
  info(message: string, context?: string): void {
    this.log("info", message, { context });
  }

  /**
   * 로그를 파일로 다운로드
   */
  downloadLogs(): void {
    if (this.logs.length === 0) {
      alert("저장된 로그가 없습니다.");
      return;
    }

    const content = this.logs
      .map((log) => {
        let logText = `[${log.timestamp}] [${log.level.toUpperCase()}] ${log.message}`;
        if (log.context) {
          logText += `\n컨텍스트: ${log.context}`;
        }
        if (log.error) {
          logText += `\n에러 이름: ${log.error.name || "N/A"}`;
          logText += `\n에러 메시지: ${log.error.message || "N/A"}`;
          logText += `\n에러 코드: ${log.error.code || "N/A"}`;
          if (log.error.stack) {
            logText += `\n스택 트레이스:\n${log.error.stack}`;
          }
        }
        logText += `\n사용자 에이전트: ${log.userAgent}`;
        logText += `\nOffice 호스트: ${log.officeHost}`;
        logText += `\nOffice 버전: ${log.officeVersion}`;
        logText += "\n" + "=".repeat(80);
        return logText;
      })
      .join("\n\n");

    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `addin-logs-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  /**
   * 서버로 로그 전송
   */
  private async sendToServer(logEntry: LogEntry): Promise<void> {
    try {
      await fetch("https://chat.k-armor.ai:8100/api/logs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(logEntry),
      });
    } catch (err) {
      // 서버 전송 실패는 조용히 무시 (무한 루프 방지)
      console.warn("로그 서버 전송 실패 (무시됨):", err);
    }
  }

  /**
   * 모든 로그를 서버로 전송
   */
  async sendAllLogsToServer(): Promise<void> {
    try {
      await fetch("https://chat.k-armor.ai:8100/api/logs/batch", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ logs: this.logs }),
      });
      this.info("모든 로그가 서버로 전송되었습니다.");
    } catch (err) {
      this.warn("로그 일괄 전송 실패", String(err));
    }
  }

  /**
   * 저장된 로그 조회
   */
  getLogs(): LogEntry[] {
    return [...this.logs];
  }

  /**
   * 에러 로그만 조회
   */
  getErrorLogs(): LogEntry[] {
    return this.logs.filter((log) => log.level === "error");
  }

  /**
   * 로그 초기화
   */
  clearLogs(): void {
    this.logs = [];
    this.info("로그가 초기화되었습니다.");
  }
}

// 전역 인스턴스 생성
export const addinLogger = new AddinLogger();

// 전역에서 접근 가능하도록 window에 추가 (디버깅용)
if (typeof window !== "undefined") {
  (window as any).addinLogger = addinLogger;
}
