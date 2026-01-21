import React, { useState } from "react";
import Sidebar, { SidebarType, WordSidebarType, ExcelSidebarType } from "../Sidebar";

const App: React.FC = () => {
  const [selectedItem, setSelectedItem] = useState<SidebarType | WordSidebarType | null>(null);

  const handleItemSelect = (item: SidebarType | WordSidebarType | ExcelSidebarType) => {
    if (item === SidebarType.Chat || item === SidebarType.Audit || item === SidebarType.Comment || item === WordSidebarType.WordDemo) {
      setSelectedItem(item);
    }
  };

  const renderContent = () => {
    if (!selectedItem) {
      return (
        <div style={{ padding: "20px", textAlign: "center", color: "#666" }}>
          <h2>Word API Tester</h2>
          <p>오른쪽 메뉴에서 항목을 선택하세요.</p>
        </div>
      );
    }

    switch (selectedItem) {
      case SidebarType.Chat:
        return <div style={{ padding: "20px" }}><h3>채팅 기능</h3><p>채팅 관련 기능이 여기에 표시됩니다.</p></div>;
      case SidebarType.Audit:
        return <div style={{ padding: "20px" }}><h3>검사 기능</h3><p>검사 관련 기능이 여기에 표시됩니다.</p></div>;
      case SidebarType.Comment:
        return <div style={{ padding: "20px" }}><h3>주석 기능</h3><p>주석 관련 기능이 여기에 표시됩니다.</p></div>;
      case WordSidebarType.WordDemo:
        return <div style={{ padding: "20px" }}><h3>Word 데모</h3><p>Word 전용 데모 기능이 여기에 표시됩니다.</p></div>;
      default:
        return <div style={{ padding: "20px" }}>알 수 없는 항목입니다.</div>;
    }
  };

  return (
    <div style={{ height: "100vh", position: "relative" }}>
      <div style={{ marginRight: "70px", height: "100%", overflowY: "auto" }}>
        {renderContent()}
      </div>
      <Sidebar
        commonItems={[SidebarType.Chat, SidebarType.Audit, SidebarType.Comment]}
        hostItems={[WordSidebarType.WordDemo]}
        onItemSelect={handleItemSelect}
        selectedItem={selectedItem}
      />
    </div>
  );
};

export default App;
