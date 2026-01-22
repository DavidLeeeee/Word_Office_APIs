import React, { useState } from "react";
import Sidebar, { SidebarType, WordSidebarType, ExcelSidebarType } from "../Sidebar";
import Chat from "../Common/Page/Chats/Chat";
import Selection from "./Selection/Selection";

const App: React.FC = () => {
  const [selectedItem, setSelectedItem] = useState<SidebarType | ExcelSidebarType | null>(null);

  const handleItemSelect = (item: SidebarType | WordSidebarType | ExcelSidebarType) => {
    if (item === SidebarType.Chat || 
        item === ExcelSidebarType.Workbook || 
        item === ExcelSidebarType.Worksheet || 
        item === ExcelSidebarType.Range || 
        item === ExcelSidebarType.Format || 
        item === ExcelSidebarType.Table || 
        item === ExcelSidebarType.Chart || 
        item === ExcelSidebarType.Pivot || 
        item === ExcelSidebarType.Validation || 
        item === ExcelSidebarType.Slicer || 
        item === ExcelSidebarType.Shapes || 
        item === ExcelSidebarType.Settings || 
        item === ExcelSidebarType.Queries || 
        item === ExcelSidebarType.Functions || 
        item === ExcelSidebarType.Audit || 
        item === ExcelSidebarType.Comment) {
      setSelectedItem(item);
    }
  };

  const renderContent = () => {
    switch (selectedItem) {
      case SidebarType.Chat:
        return <Chat />;
      case ExcelSidebarType.Range:
        return <Selection />;
      case ExcelSidebarType.Workbook:
        return <div style={{ padding: "20px" }}><h3>Excel 워크북</h3><p>워크북 관련 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Worksheet:
        return <div style={{ padding: "20px" }}><h3>Excel 워크시트</h3><p>워크시트 관련 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Format:
        return <div style={{ padding: "20px" }}><h3>Excel 서식/디자인</h3><p>셀 서식 및 디자인 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Table:
        return <div style={{ padding: "20px" }}><h3>Excel 테이블</h3><p>Excel 테이블 관련 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Chart:
        return <div style={{ padding: "20px" }}><h3>Excel 차트</h3><p>차트 생성 및 시각화 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Pivot:
        return <div style={{ padding: "20px" }}><h3>Excel 피벗</h3><p>피벗 테이블 요약/집계 자동화 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Validation:
        return <div style={{ padding: "20px" }}><h3>Excel 데이터 검증</h3><p>데이터 검증 및 규칙 강제 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Slicer:
        return <div style={{ padding: "20px" }}><h3>Excel 슬라이서</h3><p>슬라이서 필터 UX 구성 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Shapes:
        return <div style={{ padding: "20px" }}><h3>Excel 도형</h3><p>도형 및 시각 요소 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Settings:
        return <div style={{ padding: "20px" }}><h3>Excel 설정</h3><p>파일 단위 애드인 상태 저장 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Queries:
        return <div style={{ padding: "20px" }}><h3>Excel 쿼리</h3><p>워크북의 쿼리 컬렉션 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Functions:
        return <div style={{ padding: "20px" }}><h3>Excel 함수</h3><p>내장 함수 호출/연산 보조 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Audit:
        return <div style={{ padding: "20px" }}><h3>Excel 검사 기능</h3><p>Excel 전용 검사 관련 기능이 여기에 표시됩니다.</p></div>;
      case ExcelSidebarType.Comment:
        return <div style={{ padding: "20px" }}><h3>Excel 주석 기능</h3><p>Excel 전용 주석 관련 기능이 여기에 표시됩니다.</p></div>;
      default:
        return <div style={{ padding: "20px" }}>알 수 없는 항목입니다.</div>;
    }
  };

  return (
    <div style={{ height: "100vh", width: "100%", position: "relative" }}>
      <div style={{ width: "calc(100% - 70px)", height: "100%" }}>
        {renderContent()}
      </div>
      <Sidebar
        commonItems={[SidebarType.Chat]}
        hostItems={[
          ExcelSidebarType.Workbook,
          ExcelSidebarType.Worksheet,
          ExcelSidebarType.Range,
          ExcelSidebarType.Format,
          ExcelSidebarType.Table,
          ExcelSidebarType.Chart,
          ExcelSidebarType.Pivot,
          ExcelSidebarType.Validation,
          ExcelSidebarType.Slicer,
          ExcelSidebarType.Shapes,
          ExcelSidebarType.Settings,
          ExcelSidebarType.Queries,
          ExcelSidebarType.Functions,
          ExcelSidebarType.Audit,
          ExcelSidebarType.Comment
        ]}
        onItemSelect={handleItemSelect}
        selectedItem={selectedItem}
      />
    </div>
  );
};

export default App;
