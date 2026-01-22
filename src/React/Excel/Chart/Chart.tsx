import React, { useState } from "react";

/* global Excel */

const Chart: React.FC = () => {
  const [result, setResult] = useState("");
  const [chartName, setChartName] = useState("");
  const [dataAddress, setDataAddress] = useState("A1");
  const [useSelection, setUseSelection] = useState(false);
  const [chartType, setChartType] = useState<"ColumnClustered" | "Line" | "Pie" | "BarClustered" | "Area" | "XYScatter">("ColumnClustered");
  const [seriesBy, setSeriesBy] = useState<"Auto" | "Columns" | "Rows">("Auto");
  const [chartTitle, setChartTitle] = useState("");
  const [showLegend, setShowLegend] = useState(true);

  // 현재 선택된 범위 가져오기
  const getSelectedRange = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();

        if (range.address === "") {
          setResult("Excel에서 범위를 선택한 후 다시 시도해주세요.");
          return;
        }

        setDataAddress(range.address);
        setUseSelection(true);
        setResult(`선택된 범위를 가져왔습니다!\n주소: ${range.address}\n\n이제 "선택된 범위 사용" 모드가 활성화되었습니다.`);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 1. 차트 목록 가져오기
  const listCharts = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        charts.load("items/name,items/id,items/chartType,items/width,items/height");
        await context.sync();

        if (charts.items.length === 0) {
          setResult("현재 워크시트에 차트가 없습니다.");
          return;
        }

        let resultText = `차트 목록 (${charts.items.length}개):\n\n`;
        charts.items.forEach((chart, index) => {
          resultText += `${index + 1}. ${chart.name}\n`;
          resultText += `   ID: ${chart.id}\n`;
          resultText += `   타입: ${chart.chartType}\n`;
          resultText += `   크기: ${chart.width}pt × ${chart.height}pt\n\n`;
        });

        resultText += `과정:\n1. context.workbook.worksheets.getActiveWorksheet()으로 활성 시트 가져오기\n2. sheet.charts로 차트 컬렉션 가져오기\n3. charts.load("items/name,items/id,...")로 속성 로드\n4. context.sync()로 동기화`;

        setResult(resultText);
      });
    } catch (error) {
      setResult(`오류: ${error instanceof Error ? error.message : "알 수 없는 오류"}`);
    }
  };

  // 2. 차트 생성
  const createChart = async () => {
    if (!useSelection && !dataAddress.trim()) {
      setResult("데이터 범위 주소를 입력하거나 '선택된 범위 사용' 버튼을 클릭해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let range: Excel.Range;
        
        if (useSelection) {
          range = context.workbook.getSelectedRange();
        } else {
          range = sheet.getRange(dataAddress);
        }
        
        range.load("address");
        await context.sync();

        const charts = sheet.charts;
        const seriesByValue = seriesBy === "Auto" ? undefined : (seriesBy === "Columns" ? "Columns" : "Rows");
        const newChart = charts.add(chartType, range, seriesByValue);
        
        newChart.load("name,id,chartType,width,height");
        
        // 차트 제목 설정
        if (chartTitle.trim()) {
          newChart.title.text = chartTitle;
          newChart.title.visible = true;
        }
        
        // 범례 표시 설정
        newChart.legend.visible = showLegend;
        
        await context.sync();

        const method = useSelection ? "context.workbook.getSelectedRange()" : `sheet.getRange("${dataAddress}")`;
        const seriesByText = seriesBy === "Auto" ? "undefined (자동)" : seriesBy;
        setResult(`차트 생성 완료!\n데이터 범위: ${range.address}\n차트 이름: ${newChart.name}\nID: ${newChart.id}\n타입: ${newChart.chartType}\n크기: ${newChart.width}pt × ${newChart.height}pt\n시리즈 기준: ${seriesByText}\n제목: ${chartTitle || "(없음)"}\n범례 표시: ${showLegend ? "예" : "아니오"}\n\n과정:\n1. ${method}로 데이터 범위 가져오기\n2. sheet.charts.add("${chartType}", range, ${seriesByText})로 차트 생성\n3. newChart.title.text로 제목 설정 (선택)\n4. newChart.legend.visible로 범례 표시 설정\n5. context.sync()로 동기화`);
        setDataAddress("A1");
        setUseSelection(false);
        setChartTitle("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 데이터 범위가 유효하지 않거나 차트 타입과 호환되지 않는 경우 생성할 수 없습니다.`);
    }
  };

  // 3. 차트 정보 읽기
  const getChartInfo = async () => {
    if (!chartName.trim()) {
      setResult("차트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        const chart = charts.getItem(chartName);
        
        chart.load("name,id,chartType,width,height,left,top,plotBy,plotVisibleOnly,style");
        const title = chart.title;
        title.load("text,visible");
        const legend = chart.legend;
        legend.load("visible");
        await context.sync();

        const info = `차트 정보:\n\n이름: ${chart.name}\nID: ${chart.id}\n타입: ${chart.chartType}\n위치: (${chart.left}pt, ${chart.top}pt)\n크기: ${chart.width}pt × ${chart.height}pt\n시리즈 기준: ${chart.plotBy}\n보이는 셀만: ${chart.plotVisibleOnly ? "예" : "아니오"}\n스타일: ${chart.style}\n제목: ${title.text || "(없음)"}\n제목 표시: ${title.visible ? "예" : "아니오"}\n범례 표시: ${legend.visible ? "예" : "아니오"}\n\n과정:\n1. sheet.charts.getItem("${chartName}")로 차트 가져오기\n2. chart.load()로 속성 로드\n3. chart.title, chart.legend로 제목/범례 정보 로드\n4. context.sync()로 동기화`;

        setResult(info);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}\n\n참고: 차트 이름을 확인해주세요.`);
    }
  };

  // 4. 차트 타입 변경
  const changeChartType = async () => {
    if (!chartName.trim()) {
      setResult("차트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        const chart = charts.getItem(chartName);
        chart.load("name,chartType");
        await context.sync();

        const oldType = chart.chartType;
        chart.chartType = chartType;
        await context.sync();

        setResult(`차트 타입 변경 완료!\n차트: ${chart.name}\n이전 타입: ${oldType}\n새 타입: ${chart.chartType}\n\n과정:\n1. sheet.charts.getItem("${chartName}")로 차트 가져오기\n2. chart.chartType = "${chartType}"로 타입 변경\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 5. 차트 제목 설정
  const applyChartTitle = async () => {
    if (!chartName.trim()) {
      setResult("차트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        const chart = charts.getItem(chartName);
        chart.load("name");
        const title = chart.title;
        title.load("text");
        await context.sync();

        const oldTitle = title.text;
        title.text = chartTitle.trim() || "";
        title.visible = chartTitle.trim() !== "";
        await context.sync();

        setResult(`차트 제목 설정 완료!\n차트: ${chart.name}\n이전 제목: ${oldTitle || "(없음)"}\n새 제목: ${title.text || "(없음)"}\n표시: ${title.visible ? "예" : "아니오"}\n\n과정:\n1. sheet.charts.getItem("${chartName}")로 차트 가져오기\n2. chart.title.text = "${chartTitle}"로 제목 설정\n3. chart.title.visible로 표시 여부 설정\n4. context.sync()로 동기화`);
        setChartTitle("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 6. 범례 표시 설정
  const setLegendVisibility = async () => {
    if (!chartName.trim()) {
      setResult("차트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        const chart = charts.getItem(chartName);
        chart.load("name");
        const legend = chart.legend;
        legend.load("visible");
        await context.sync();

        const oldVisibility = legend.visible;
        legend.visible = showLegend;
        await context.sync();

        setResult(`범례 표시 설정 완료!\n차트: ${chart.name}\n이전 표시: ${oldVisibility ? "예" : "아니오"}\n새 표시: ${legend.visible ? "예" : "아니오"}\n\n과정:\n1. sheet.charts.getItem("${chartName}")로 차트 가져오기\n2. chart.legend.visible = ${showLegend}로 범례 표시 설정\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 7. 차트 삭제
  const deleteChart = async () => {
    if (!chartName.trim()) {
      setResult("삭제할 차트 이름을 입력해주세요.");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        const chart = charts.getItem(chartName);
        chart.load("name");
        await context.sync();

        const deletedName = chart.name;
        chart.delete();
        await context.sync();

        setResult(`차트 삭제 완료!\n삭제된 차트: ${deletedName}\n\n과정:\n1. sheet.charts.getItem("${chartName}")로 차트 가져오기\n2. chart.delete()로 차트 삭제\n3. context.sync()로 동기화`);
        setChartName("");
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  // 8. 차트 크기 설정
  const setChartSize = async () => {
    if (!chartName.trim()) {
      setResult("차트 이름을 입력해주세요.");
      return;
    }

    const width = parseFloat((document.getElementById("chartWidth") as HTMLInputElement)?.value || "400");
    const height = parseFloat((document.getElementById("chartHeight") as HTMLInputElement)?.value || "300");

    if (isNaN(width) || isNaN(height) || width <= 0 || height <= 0) {
      setResult("유효한 너비와 높이를 입력해주세요 (양수).");
      return;
    }

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        const chart = charts.getItem(chartName);
        chart.load("name,width,height");
        await context.sync();

        const oldWidth = chart.width;
        const oldHeight = chart.height;
        chart.width = width;
        chart.height = height;
        await context.sync();

        setResult(`차트 크기 설정 완료!\n차트: ${chart.name}\n이전 크기: ${oldWidth}pt × ${oldHeight}pt\n새 크기: ${chart.width}pt × ${chart.height}pt\n\n과정:\n1. sheet.charts.getItem("${chartName}")로 차트 가져오기\n2. chart.width = ${width}, chart.height = ${height}로 크기 설정\n3. context.sync()로 동기화`);
      });
    } catch (error: any) {
      setResult(`오류 발생!\n\n오류 코드: ${error.code || "알 수 없음"}\n오류 메시지: ${error.message}`);
    }
  };

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", overflow: "hidden" }}>
      <div style={{ padding: "15px", borderBottom: "1px solid #ddd", backgroundColor: "#f5f5f5", overflowY: "auto", flex: "0 0 auto", maxHeight: "60%" }}>
        <h3 style={{ margin: "0 0 15px 0" }}>Excel 차트</h3>

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
          <h4 style={{ margin: "0 0 10px 0", color: "#1976d2", fontSize: "14px" }}>📊 Excel 차트 안내</h4>
          <p style={{ margin: "0 0 8px 0", color: "#1976d2" }}>
            Excel 차트는 데이터를 시각적으로 표현하는 강력한 기능입니다. 다양한 차트 타입과 스타일을 지원합니다.
          </p>
          <p style={{ margin: "8px 0", color: "#1976d2", fontWeight: "bold" }}>✅ 지원되는 기능:</p>
          <ul style={{ margin: "0 0 8px 0", paddingLeft: "20px", color: "#1976d2" }}>
            <li>차트 생성 (다양한 차트 타입)</li>
            <li>차트 목록 조회</li>
            <li>차트 정보 읽기</li>
            <li>차트 타입 변경</li>
            <li>차트 제목 설정</li>
            <li>범례 표시 설정</li>
            <li>차트 크기 조정</li>
            <li>차트 삭제</li>
          </ul>
        </div>

        {/* 차트 목록 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #4caf50" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#4caf50" }}>📋 차트 목록</h4>
          <button
            onClick={listCharts}
            style={{
              padding: "8px 16px",
              backgroundColor: "#4caf50",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            차트 목록 보기
          </button>
        </div>

        {/* 차트 생성 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #ff9800" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#ff9800" }}>➕ 차트 생성</h4>
          <div style={{ display: "flex", gap: "10px", marginBottom: "10px", alignItems: "center" }}>
            <button
              onClick={getSelectedRange}
              style={{
                padding: "8px 16px",
                backgroundColor: useSelection ? "#4caf50" : "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontWeight: useSelection ? "bold" : "normal",
              }}
            >
              {useSelection ? "✓ 선택된 범위 사용 중" : "선택된 범위 사용"}
            </button>
            <button
              onClick={() => {
                setUseSelection(false);
                setResult("직접 입력 모드로 전환되었습니다.");
              }}
              style={{
                padding: "8px 16px",
                backgroundColor: !useSelection ? "#4caf50" : "#2196f3",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontWeight: !useSelection ? "bold" : "normal",
              }}
            >
              직접 입력
            </button>
          </div>
          <input
            type="text"
            value={dataAddress}
            onChange={(e) => {
              setDataAddress(e.target.value);
              setUseSelection(false);
            }}
            placeholder={useSelection ? "선택된 범위 사용 중..." : "예: A1:C10"}
            disabled={useSelection}
            style={{
              width: "100%",
              padding: "8px",
              border: "1px solid #ddd",
              borderRadius: "5px",
              marginBottom: "10px",
              backgroundColor: useSelection ? "#f5f5f5" : "#fff",
              cursor: useSelection ? "not-allowed" : "text",
            }}
          />
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>차트 타입:</label>
            <select
              value={chartType}
              onChange={(e) => setChartType(e.target.value as any)}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            >
              <option value="ColumnClustered">세로 막대형 (묶은 세로 막대형)</option>
              <option value="Line">꺾은선형</option>
              <option value="Pie">원형</option>
              <option value="BarClustered">가로 막대형 (묶은 가로 막대형)</option>
              <option value="Area">영역형</option>
              <option value="XYScatter">분산형 (XY)</option>
            </select>
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "block", fontSize: "12px", marginBottom: "5px" }}>시리즈 기준:</label>
            <select
              value={seriesBy}
              onChange={(e) => setSeriesBy(e.target.value as any)}
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            >
              <option value="Auto">자동</option>
              <option value="Columns">열</option>
              <option value="Rows">행</option>
            </select>
          </div>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={chartTitle}
              onChange={(e) => setChartTitle(e.target.value)}
              placeholder="차트 제목 (선택사항)"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ marginBottom: "10px" }}>
            <label style={{ display: "flex", alignItems: "center", gap: "10px", cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={showLegend}
                onChange={(e) => setShowLegend(e.target.checked)}
              />
              <span>범례 표시</span>
            </label>
          </div>
          <button
            onClick={createChart}
            style={{
              padding: "8px 16px",
              backgroundColor: "#ff9800",
              color: "#fff",
              border: "none",
              borderRadius: "5px",
              cursor: "pointer",
            }}
          >
            차트 생성
          </button>
        </div>

        {/* 차트 조작 */}
        <div style={{ marginBottom: "20px", padding: "15px", backgroundColor: "#fff", borderRadius: "5px", border: "1px solid #9c27b0" }}>
          <h4 style={{ margin: "0 0 10px 0", color: "#9c27b0" }}>🔧 차트 조작</h4>
          <div style={{ marginBottom: "10px" }}>
            <input
              type="text"
              value={chartName}
              onChange={(e) => setChartName(e.target.value)}
              placeholder="차트 이름"
              style={{ width: "100%", padding: "8px", border: "1px solid #ddd", borderRadius: "5px", marginBottom: "10px" }}
            />
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "10px" }}>
            <button
              onClick={getChartInfo}
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
              onClick={changeChartType}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              타입 변경
            </button>
            <button
              onClick={applyChartTitle}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              제목 설정
            </button>
            <button
              onClick={setLegendVisibility}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              범례 표시 설정
            </button>
            <button
              onClick={deleteChart}
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
          <div style={{ display: "flex", gap: "10px", alignItems: "center", marginBottom: "10px" }}>
            <input
              id="chartWidth"
              type="number"
              placeholder="너비 (pt)"
              defaultValue="400"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <input
              id="chartHeight"
              type="number"
              placeholder="높이 (pt)"
              defaultValue="300"
              style={{ flex: 1, padding: "8px", border: "1px solid #ddd", borderRadius: "5px" }}
            />
            <button
              onClick={setChartSize}
              style={{
                padding: "8px 16px",
                backgroundColor: "#9c27b0",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
              }}
            >
              크기 설정
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
          {result || "위 버튼을 클릭하여 Excel 차트 기능을 테스트해보세요."}
        </pre>
      </div>
    </div>
  );
};

export default Chart;
