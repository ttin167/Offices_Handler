/**
 * src/modules/constants.js — Standard Constants & Presets for Antigravity Office Studio
 */

export const DIAGRAM_PRESETS = {
  mermaid_sequence: `sequenceDiagram
  autonumber
  actor User as Khách Hàng
  participant Web as Web Studio (Vite)
  participant API as Python API Server
  participant Engine as Office Engine (OpenXML)

  User->>Web: Tải lên tài liệu<br/>PDF / DOCX
  Web->>API: POST /api/upload<br/>(Multipart)
  API-->>Web: Trả về file_path<br/>& metadata
  Web->>API: POST /api/pdf-to-pptx<br/>(Theme: Thesis Blue)
  API->>Engine: Bóc tách AST &<br/>Render 16:9 Slide
  Engine-->>API: Trả về tệp<br/>presentation.pptx
  API-->>Web: 200 OK + Spec JSON<br/>& Tải Tệp
  Web-->>User: Hiển thị bộ thẻ Slide<br/>& Nút Tải Về`,

  plantuml_c4: `@startuml
!include <C4/C4_Context>
title Sơ Đồ Kiến Trúc Hệ Thống (C4 Context)

Person(user, "Chuyên Viên / Giảng Viên", "Người soạn thảo tài liệu kỹ thuật & slide thuyết trình")
System(studio, "Antigravity Office Studio", "Nền tảng chuyển đổi tài liệu & biên tập slide đa năng")
System_Ext(ms_office, "Microsoft Office / 365", "Đọc và trình chiếu tệp PPTX, DOCX chuẩn OpenXML")
System_Ext(pdf_reader, "Trình Đọc PDF", "Hiển thị tài liệu phân giải cao 300+ DPI")

Rel(user, studio, "Tải lên DOCX/PDF & Tùy chọn Theme")
Rel(studio, ms_office, "Xuất bản tệp 16:9 Zero-Overflow")
Rel(studio, pdf_reader, "Xuất bản PDF chất lượng cao")
@enduml`,

  mermaid_erd: `erDiagram
  DOCUMENT ||--o{ SLIDE : contains
  SLIDE ||--o{ BLOCK : composed_of
  BLOCK ||--o| TABLE_DATA : renders
  DOCUMENT {
    string id PK
    string filename
    string original_format
    int page_count
  }
  SLIDE {
    int slide_index PK
    string title
    string layout_type
    string background_color
  }
  BLOCK {
    string block_id PK
    string block_type
    string content
  }
  TABLE_DATA {
    string table_id PK
    int row_count
    int col_count
  }`,

  mermaid_flowchart: `flowchart TD
  Start([Bắt Đầu Tiến Trình]) --> Step1[Nhận Yêu Cầu Xử Lý Tài Liệu]
  Step1 --> Check{"Kiểm Tra Hợp Lệ?"}
  Check -->|Hợp Lệ| Process[Tiến Hành Bóc Tách OpenXML AST]
  Check -->|Không Hợp Lệ| Error[Thông Báo Lỗi Định Dạng]
  Process --> DB[(Ghi Dữ Liệu Vào DB)]
  DB --> Success([Hoàn Tất & Xuất Bản 300+ DPI])
  Error --> EndFail([Kết Thúc Thất Bại])`,

  plantuml_usecase: `@startuml
left to right direction
actor "Khách Hàng" as User
actor "Quản Trị Viên" as Admin

rectangle "Hệ Thống Office Studio" {
  usecase "Tải Lên Tài Liệu" as UC1
  usecase "Chuyển Đổi DOCX sang PPTX" as UC2
  usecase "Kiểm Tra Lỗi QA 12 Cổng" as UC3
  usecase "Cấu Hình Theme Trình Bày" as UC4
  usecase "Quản Lý Người Dùng" as UC5
}

User --> UC1
User --> UC2
User --> UC3
UC2 ..> UC4 : <<include>>
Admin --> UC3
Admin --> UC5
@enduml`,

  plantuml_swimlane: `@startuml
|Người Dùng|
start
:Tải lên tệp Word/PDF;

|Hệ Thống Web Studio|
:Tiếp nhận và lưu trữ tạm thời;
:Gửi yêu cầu phân tích tới API;

|Bộ Máy Xử Lý Office Engine|
:Phân tích cấu trúc bảng & styles;
if (Bảng vượt quá lề in?) then (Có)
  :Tự động căn chỉnh độ rộng cột;
else (Không)
  :Giữ nguyên bố cục;
endif
:Tạo tài liệu PPTX 16:9;

|Hệ Thống Web Studio|
:Trả về đường dẫn tải tệp;

|Người Dùng|
:Tải về và kiểm tra tài liệu;
stop
@enduml`,

  mermaid_state: `stateDiagram-v2
  [*] --> Draft : Tạo mới tài liệu
  Draft --> InReview : Nộp bản thảo
  InReview --> Approved : Duyệt hợp lệ
  InReview --> Rejected : Yêu cầu sửa đổi
  Rejected --> Draft : Chỉnh sửa nội dung
  Approved --> Published : Xuất bản 300 DPI
  Published --> Archived : Lưu trữ hệ thống
  Archived --> [*]`,

  plantuml_class: `@startuml
class DocumentNode {
  +String id
  +int pageIndex
  +render(): Image
}

class SlideModel {
  +String title
  +String themeName
  +List<CardDef> cards
  +buildPresentation(): void
}

class TableModel {
  +int rowCount
  +int colCount
  +applyOpenXmlInvariants(): void
}

DocumentNode <|-- SlideModel
DocumentNode <|-- TableModel
@enduml`,

  mermaid_mindmap: `mindmap
  root((Office Studio))
    Document Processing
      PDF Lossless
      DOCX Repair
      Markdown AST
    Technical Diagrams
      Mermaid
      PlantUML C4
      Draw.io mxGraph
      SVG Canvas Spec
    Presentations
      16:9 Widescreen
      Thesis Themes
      Auto-Card Layout`,

  canvas_flow: JSON.stringify({
    "canvas": { "width": 1400, "height": 770 },
    "nodes": [
      { "id": "user", "title": "Client Actor", "type": "actor", "x": 80, "y": 280, "width": 70, "height": 100, "icon": "user" },
      { "id": "gateway", "title": "API Gateway", "type": "primary", "x": 260, "y": 290, "width": 180, "height": 80, "icon": "shield" },
      { "id": "decision", "title": "Auth Valid?", "type": "decision", "x": 540, "y": 280, "width": 140, "height": 100 },
      { "id": "cloud_svc", "title": "External Cloud AI", "type": "cloud", "x": 780, "y": 120, "width": 180, "height": 90, "icon": "cloud" },
      { "id": "doc_worker", "title": "Document Worker", "type": "component", "x": 780, "y": 300, "width": 180, "height": 80, "icon": "server" },
      { "id": "msg_queue", "title": "Async Task Queue", "type": "queue", "x": 780, "y": 480, "width": 180, "height": 70 },
      { "id": "database", "title": "Primary Database", "type": "database", "x": 1080, "y": 300, "width": 140, "height": 90, "icon": "database" },
      { "id": "finish", "title": "Completed", "type": "terminator", "x": 1080, "y": 490, "width": 140, "height": 60, "icon": "zap" }
    ],
    "edges": [
      { "source": "user", "target": "gateway", "label": "Gửi Yêu Cầu", "exit_port": "right", "entry_port": "left" },
      { "source": "gateway", "target": "decision", "label": "Xác Thực JWT", "exit_port": "right", "entry_port": "left" },
      { "source": "decision", "target": "cloud_svc", "label": "AI Prompt", "exit_port": "top", "entry_port": "left" },
      { "source": "decision", "target": "doc_worker", "label": "Hợp Lệ", "exit_port": "right", "entry_port": "left" },
      { "source": "decision", "target": "msg_queue", "label": "Hàng Đợi Batch", "exit_port": "bottom", "entry_port": "left" },
      { "source": "doc_worker", "target": "database", "label": "Ghi AST OpenXML", "exit_port": "right", "entry_port": "left" },
      { "source": "msg_queue", "target": "finish", "label": "Hoàn Tất Xử Lý", "exit_port": "right", "entry_port": "left" }
    ]
  }, null, 2),

  drawio_overview: `<mxfile host="Antigravity Studio" version="21.0.0">
  <diagram id="arch" name="Enterprise Architecture">
    <mxGraphModel dx="1400" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1400" pageHeight="800" background="#FFFFFF">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="actor1" value="End User" style="shape=umlActor;verticalLabelPosition=bottom;verticalAlign=top;html=1;outlineConnect=0;fillColor=#FFFFFF;strokeColor=#1E293B;strokeWidth=1.8;" vertex="1" parent="1">
          <mxGeometry x="120" y="240" width="60" height="90" as="geometry"/>
        </mxCell>
        <mxCell id="hub" value="Antigravity Gateway Hub" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#0051E2;strokeColor=#003DB3;fontColor=#FFFFFF;fontStyle=1;fontSize=14;" vertex="1" parent="1">
          <mxGeometry x="320" y="250" width="220" height="70" as="geometry"/>
        </mxCell>
        <mxCell id="dec1" value="Is Valid?" style="rhombus;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#D97706;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">
          <mxGeometry x="640" y="240" width="130" height="90" as="geometry"/>
        </mxCell>
        <mxCell id="cloud1" value="External AI Cloud" style="ellipse;shape=cloud;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#0284C7;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">
          <mxGeometry x="900" y="110" width="180" height="95" as="geometry"/>
        </mxCell>
        <mxCell id="db1" value="PostgreSQL DB" style="shape=cylinder3;whiteSpace=wrap;html=1;boundedLbl=1;backgroundOutline=1;size=15;fillColor=#F8FAFC;strokeColor=#2563EB;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">
          <mxGeometry x="920" y="360" width="140" height="100" as="geometry"/>
        </mxCell>
        <mxCell id="e1" edge="1" parent="1" source="actor1" target="hub" style="edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=1.8;strokeColor=#475569;exitX=1;exitY=0.5;entryX=0;entryY=0.5;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="e2" edge="1" parent="1" source="hub" target="dec1" style="edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=1.8;strokeColor=#475569;exitX=1;exitY=0.5;entryX=0;entryY=0.5;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="e3" value="AI Request" edge="1" parent="1" source="dec1" target="cloud1" style="edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=1.8;strokeColor=#0284C7;exitX=0.5;exitY=0;entryX=0;entryY=0.5;labelBackgroundColor=#ffffff;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="e4" value="Store AST" edge="1" parent="1" source="dec1" target="db1" style="edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeWidth=1.8;strokeColor=#2563EB;exitX=0.5;exitY=1;entryX=0;entryY=0.5;labelBackgroundColor=#ffffff;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>`,
};
