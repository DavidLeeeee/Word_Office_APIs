import React, { useState } from "react";
import Sidebar, { SidebarType, WordSidebarType, ExcelSidebarType } from "../Sidebar";
import Chat from "../Common/Page/Chats/Chat";
import Workbook from "./Workbook/Workbook";
import Worksheet from "./Worksheet/Worksheet";
import Range from "./Range/Range";
import Format from "./Format/Format";
import Table from "./Table/Table";
import Chart from "./Chart/Chart";
import Pivot from "./Pivot/Pivot";
import Validation from "./Validation/Validation";
import Slicer from "./Slicer/Slicer";
import Shapes from "./Shapes/Shapes";
import Settings from "./Settings/Settings";
import Queries from "./Queries/Queries";
import Functions from "./Functions/Functions";
import Audit from "./Audit/Audit";
import Comment from "./Comment/Comment";
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
      case ExcelSidebarType.Workbook:
        return <Workbook />;
      case ExcelSidebarType.Worksheet:
        return <Worksheet />;
      case ExcelSidebarType.Range:
        return <Range />;
      case ExcelSidebarType.Format:
        return <Format />;
      case ExcelSidebarType.Table:
        return <Table />;
      case ExcelSidebarType.Chart:
        return <Chart />;
      case ExcelSidebarType.Pivot:
        return <Pivot />;
      case ExcelSidebarType.Validation:
        return <Validation />;
      case ExcelSidebarType.Slicer:
        return <Slicer />;
      case ExcelSidebarType.Shapes:
        return <Shapes />;
      case ExcelSidebarType.Settings:
        return <Settings />;
      case ExcelSidebarType.Queries:
        return <Queries />;
      case ExcelSidebarType.Functions:
        return <Functions />;
      case ExcelSidebarType.Audit:
        return <Audit />;
      case ExcelSidebarType.Comment:
        return <Comment />;
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
