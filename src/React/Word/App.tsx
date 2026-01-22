import React, { useState } from "react";
import Sidebar, { SidebarType, WordSidebarType, ExcelSidebarType } from "../Sidebar";
import Chat from "../Common/Page/Chats/Chat";
import Editor from "./Editor/Editor";
import Templates_builder from "./Templates_builder/Templates_builder";
import Contents from "./Contents/Contents";
import Shapes from "./Shapes/Shapes";
import Comment from "./Comment/Comment";
import Metadata from "./Metadata/Metadata";
import Format from "./Format/Format";

const App: React.FC = () => {
  const [selectedItem, setSelectedItem] = useState<SidebarType | WordSidebarType | null>(null);

  const handleItemSelect = (item: SidebarType | WordSidebarType | ExcelSidebarType) => {
    if (item === SidebarType.Chat || item === WordSidebarType.Audit || item === WordSidebarType.Comment || item === WordSidebarType.Edit || item === WordSidebarType.Template || item === WordSidebarType.Contents || item === WordSidebarType.Shapes || item === WordSidebarType.Metadata || item === WordSidebarType.Format) {
      setSelectedItem(item);
    }
  };

  const renderContent = () => {


    switch (selectedItem) {
      case SidebarType.Chat:
        return <Chat />;
      case WordSidebarType.Audit:
        return <div style={{ padding: "20px" }}><h3>Word 검사 기능</h3><p>Word 전용 검사 관련 기능이 여기에 표시됩니다.</p></div>;
      case WordSidebarType.Comment:
        return <Comment />;
      case WordSidebarType.Edit:
        return <Editor />;
      case WordSidebarType.Template:
        return <Templates_builder />;
      case WordSidebarType.Contents:
        return <Contents />;
      case WordSidebarType.Shapes:
        return <Shapes />;
      case WordSidebarType.Metadata:
        return <Metadata />;
      case WordSidebarType.Format:
        return <Format />;
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
        hostItems={[WordSidebarType.Edit, WordSidebarType.Template, WordSidebarType.Contents, WordSidebarType.Shapes, WordSidebarType.Audit, WordSidebarType.Comment, WordSidebarType.Metadata, WordSidebarType.Format]}
        onItemSelect={handleItemSelect}
        selectedItem={selectedItem}
      />
    </div>
  );
};

export default App;
