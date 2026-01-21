// 재사용 가능한 Sidebar 컴포넌트를 만든다.

// Word, Excel에서 해당 Sidebar를 불러 렌더링을 한다.

// 공용 사이드바 항목
enum SidebarType {
    Chat = "채팅",
    Audit = "검사",
    Comment = "주석",
}

// Word 사이드바 항목
enum WordSidebarType {
    WordDemo = "Word 데모",
}

// Excel 사이드바 항목
enum ExcelSidebarType {
    ExcelDemo = "Excel 데모",
}

// 위 enum들에 대해서, 각각의 App.tsx에서 parameter로 받을 수 있다.
// Sidebar.tsx에서는 해당 parameter에 따라서 사이드바를 표시해주고, 메인 화면에 항목별 렌더링을 유도하는 역할을 한다.