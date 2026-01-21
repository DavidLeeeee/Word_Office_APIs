import React, { useState } from "react";
import Sidebar, { SidebarType, WordSidebarType, ExcelSidebarType } from "../Sidebar";
import Chat from "../Common/Page/Chats/Chat";

const App: React.FC = () => {
  const [selectedItem, setSelectedItem] = useState<SidebarType | ExcelSidebarType | null>(null);

  const handleItemSelect = (item: SidebarType | WordSidebarType | ExcelSidebarType) => {
    if (item === SidebarType.Chat || item === ExcelSidebarType.Audit || item === ExcelSidebarType.Comment) {
      setSelectedItem(item);
    }
  };

  const renderContent = () => {


    switch (selectedItem) {
      case SidebarType.Chat:
        return <Chat />;
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
        hostItems={[ExcelSidebarType.Audit, ExcelSidebarType.Comment]}
        onItemSelect={handleItemSelect}
        selectedItem={selectedItem}
      />
    </div>
  );
};

export default App;
