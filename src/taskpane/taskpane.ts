/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import React from "react";
import { createRoot } from "react-dom/client";
import WordApp from "../React/Word/App";
import ExcelApp from "../React/Excel/App";
import { addinLogger } from "../React/Common/Function/AddinLogger";
import { ErrorBoundary } from "../React/Common/Function/ErrorBoundary";

// 전역 에러 핸들러 설정
window.addEventListener("error", (event) => {
  addinLogger.error("전역 JavaScript 에러", event.error, "GlobalErrorHandler");
});

window.addEventListener("unhandledrejection", (event) => {
  addinLogger.error("처리되지 않은 Promise 거부", event.reason, "UnhandledRejection");
});

// Office.js 초기화 전 에러 캡처
try {
  Office.onReady((info) => {
    try {
      addinLogger.info("Office.onReady 완료", {
        context: `Host: ${info.host}, Platform: ${info.platform}`,
      });

      document.getElementById("sideload-msg")!.style.display = "none";
      document.getElementById("app-body")!.style.display = "flex";

      const container = document.getElementById("react-app");
      if (!container) {
        throw new Error("react-app 컨테이너를 찾을 수 없습니다.");
      }

      const root = createRoot(container);

      // ErrorBoundary로 감싸서 렌더링
      let AppComponent: React.ComponentType;
      if (info.host === Office.HostType.Word) {
        AppComponent = WordApp;
        addinLogger.info("Word App 렌더링 시작");
      } else if (info.host === Office.HostType.Excel) {
        AppComponent = ExcelApp;
        addinLogger.info("Excel App 렌더링 시작");
      } else {
        const UnsupportedHost = () => {
          return React.createElement(
            "div",
            { style: { padding: "20px" } },
            `지원하지 않는 호스트: ${info.host}`
          );
        };
        AppComponent = UnsupportedHost;
        addinLogger.warn(`지원하지 않는 호스트: ${info.host}`);
      }

      root.render(
        React.createElement(
          ErrorBoundary,
          null,
          React.createElement(AppComponent)
        )
      );

      addinLogger.info("Add-in 초기화 완료");
    } catch (error) {
      addinLogger.error("Office.onReady 콜백 내부 에러", error, "Office.onReady");
      
      // 에러 UI 표시
      const errorContainer = document.getElementById("react-app");
      if (errorContainer) {
        errorContainer.innerHTML = `
          <div style="padding: 20px; background-color: #fff3cd; border: 2px solid #ffc107; border-radius: 5px; margin: 20px;">
            <h2 style="color: #856404; margin-top: 0;">⚠️ Add-in 초기화 오류</h2>
            <p style="color: #856404;">Add-in을 초기화하는 중 오류가 발생했습니다.</p>
            <pre style="background-color: #fff; padding: 10px; border-radius: 4px; overflow: auto; font-size: 12px; color: #d32f2f;">
              ${error instanceof Error ? error.toString() : String(error)}
            </pre>
            <button onclick="window.addinLogger.downloadLogs()" style="padding: 10px 20px; background-color: #ffc107; color: #000; border: none; border-radius: 4px; cursor: pointer; margin-top: 15px; font-weight: bold;">
              로그 다운로드
            </button>
            <button onclick="window.location.reload()" style="padding: 10px 20px; background-color: #1976d2; color: #fff; border: none; border-radius: 4px; cursor: pointer; margin-top: 15px; margin-left: 10px; font-weight: bold;">
              페이지 새로고침
            </button>
          </div>
        `;
      }
    }
  });
} catch (error) {
  // Office.onReady 자체가 실패한 경우
  addinLogger.error("Office.onReady 호출 실패", error, "Office.onReady");
  
  const errorContainer = document.getElementById("react-app") || document.body;
  errorContainer.innerHTML = `
    <div style="padding: 20px; background-color: #f8d7da; border: 2px solid #dc3545; border-radius: 5px; margin: 20px;">
      <h2 style="color: #721c24; margin-top: 0;">❌ 치명적 오류</h2>
      <p style="color: #721c24;">Office.js를 초기화할 수 없습니다.</p>
      <pre style="background-color: #fff; padding: 10px; border-radius: 4px; overflow: auto; font-size: 12px; color: #d32f2f;">
        ${error instanceof Error ? error.toString() : String(error)}
      </pre>
      <p style="color: #721c24; margin-top: 15px;">
        가능한 원인:<br>
        1. Office.js가 로드되지 않았습니다<br>
        2. Add-in이 올바르게 sideload되지 않았습니다<br>
        3. 브라우저 호환성 문제가 있습니다
      </p>
      <button onclick="window.addinLogger.downloadLogs()" style="padding: 10px 20px; background-color: #dc3545; color: #fff; border: none; border-radius: 4px; cursor: pointer; margin-top: 15px; font-weight: bold;">
        로그 다운로드
      </button>
    </div>
  `;
}
