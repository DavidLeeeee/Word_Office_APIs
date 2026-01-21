import React, { useState } from "react";

// 공용 사이드바 항목
export enum SidebarType {
  Chat = "채팅",
}

// Word 사이드바 항목
export enum WordSidebarType {
  Audit = "검사",
  Comment = "주석",
  Edit = "수정",
  Template = "템플릿",
  Contents = "콘텐츠",
  Shapes = "도형",
  Metadata = "메타데이터",
}

// Excel 사이드바 항목
export enum ExcelSidebarType {
  Audit = "검사",
  Comment = "주석",
}

type SidebarItem = SidebarType | WordSidebarType | ExcelSidebarType;

interface SidebarProps {
  commonItems: SidebarType[];
  hostItems: SidebarItem[];
  onItemSelect: (item: SidebarItem) => void;
  selectedItem: SidebarItem | null;
}

const Sidebar: React.FC<SidebarProps> = ({ commonItems, hostItems, onItemSelect, selectedItem }) => {
  const isSelected = (item: SidebarItem) => selectedItem === item;

  return (
    <div style={{ 
      position: "fixed",
      right: 0,
      top: 0,
      width: "70px", 
      padding: "12px 6px",
      background: "linear-gradient(180deg, #1a1a2e 0%, #16213e 100%)",
      height: "100vh",
      overflowY: "auto",
      boxShadow: "-2px 0 20px rgba(0, 0, 0, 0.3)",
      zIndex: 1000
    }}>
      {commonItems.length > 0 && (
        <div style={{ marginBottom: "20px" }}>
          {commonItems.map((item) => (
            <div
              key={item}
              onClick={() => onItemSelect(item)}
              title={item}
              style={{
                padding: "10px 6px",
                cursor: "pointer",
                background: isSelected(item) 
                  ? "linear-gradient(135deg, #667eea 0%, #764ba2 100%)" 
                  : "transparent",
                borderRadius: "8px",
                marginBottom: "6px",
                fontSize: "11px",
                textAlign: "center",
                wordBreak: "break-word",
                lineHeight: "1.3",
                color: isSelected(item) ? "#ffffff" : "#a0a0b8",
                fontWeight: isSelected(item) ? "600" : "400",
                transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
                transform: isSelected(item) ? "scale(1.05)" : "scale(1)",
                boxShadow: isSelected(item) 
                  ? "0 4px 12px rgba(102, 126, 234, 0.4)" 
                  : "none"
              }}
              onMouseEnter={(e) => {
                if (!isSelected(item)) {
                  e.currentTarget.style.background = "rgba(102, 126, 234, 0.15)";
                  e.currentTarget.style.color = "#ffffff";
                  e.currentTarget.style.transform = "scale(1.02)";
                }
              }}
              onMouseLeave={(e) => {
                if (!isSelected(item)) {
                  e.currentTarget.style.background = "transparent";
                  e.currentTarget.style.color = "#a0a0b8";
                  e.currentTarget.style.transform = "scale(1)";
                }
              }}
            >
              {item}
            </div>
          ))}
        </div>
      )}

      {hostItems.length > 0 && (
        <div style={{ 
          borderTop: "1px solid rgba(160, 160, 184, 0.2)", 
          paddingTop: "15px",
          marginTop: "10px"
        }}>
          {hostItems.map((item) => (
            <div
              key={item}
              onClick={() => onItemSelect(item)}
              title={item}
              style={{
                padding: "10px 6px",
                cursor: "pointer",
                background: isSelected(item) 
                  ? "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)" 
                  : "transparent",
                borderRadius: "8px",
                marginBottom: "6px",
                fontSize: "11px",
                textAlign: "center",
                wordBreak: "break-word",
                lineHeight: "1.3",
                color: isSelected(item) ? "#ffffff" : "#a0a0b8",
                fontWeight: isSelected(item) ? "600" : "400",
                transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
                transform: isSelected(item) ? "scale(1.05)" : "scale(1)",
                boxShadow: isSelected(item) 
                  ? "0 4px 12px rgba(245, 87, 108, 0.4)" 
                  : "none"
              }}
              onMouseEnter={(e) => {
                if (!isSelected(item)) {
                  e.currentTarget.style.background = "rgba(245, 87, 108, 0.15)";
                  e.currentTarget.style.color = "#ffffff";
                  e.currentTarget.style.transform = "scale(1.02)";
                }
              }}
              onMouseLeave={(e) => {
                if (!isSelected(item)) {
                  e.currentTarget.style.background = "transparent";
                  e.currentTarget.style.color = "#a0a0b8";
                  e.currentTarget.style.transform = "scale(1)";
                }
              }}
            >
              {item}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default Sidebar;