export type DiagramEngineType = 'mermaid' | 'plantuml' | 'drawio' | 'canvas';

export interface DiagramPresetItem {
  id: string;
  label: string;
  engine: DiagramEngineType;
  code: string;
}

export const DIAGRAM_PRESETS_LIST: DiagramPresetItem[] = [
  {
    id: 'mermaid_flowchart',
    label: 'Flowchart Quy trình',
    engine: 'mermaid',
    code: `flowchart TD
  Start([Tiếp nhận Văn bản]) --> Step1[Phân loại Hồ sơ & Scan]
  Step1 --> Check{"Đủ chữ ký & dấu?"}
  Check -->|Hợp lệ| Process[Lưu trữ Số hóa PDF/A]
  Check -->|Thiếu| Reject[Trả lại bổ sung]
  Process --> Notify([Gửi thông báo hoàn tất])
  Reject --> EndFail([Chờ chỉnh sửa])`,
  },
  {
    id: 'mermaid_sequence',
    label: 'Sequence API',
    engine: 'mermaid',
    code: `sequenceDiagram
  autonumber
  actor User as Chuyên viên
  participant Studio as Office Studio Local
  participant Engine as Python OpenXML Engine

  User->>Studio: Tải lên Word / PDF
  Studio->>Engine: POST /api/docx-to-pptx
  Engine-->>Studio: Kết xuất Slide 16:9 chuẩn hóa
  Studio-->>User: Xem trước 1080p & Tải tệp`,
  },
  {
    id: 'mermaid_erd',
    label: 'ERD Dữ liệu',
    engine: 'mermaid',
    code: `erDiagram
  PHONG_BAN ||--o{ NHAN_VIEN : quan_ly
  NHAN_VIEN ||--o{ HOP_DONG : ky_ket
  PHONG_BAN {
    string ma_pb PK
    string ten_pb
  }
  NHAN_VIEN {
    string ma_nv PK
    string ho_ten
    string chuc_vu
  }
  HOP_DONG {
    string so_hd PK
    date ngay_ky
  }`,
  },
  {
    id: 'mermaid_mindmap',
    label: 'Mindmap Ý tưởng',
    engine: 'mermaid',
    code: `mindmap
  root((Tin Học Văn Phòng))
    Word Hành Chính
      Nghị định 30/2020
      Trộn thư Excel -> DOCX
      Đối chiếu Track Changes
    Excel Power Suite
      Làm sạch & Tách tên
      Pivot & Biểu đồ
      57 Công thức mẫu
    PDF Pro Studio
      Ký vẽ & Giáp lai
      Che PII & PDF/A-2b`,
  },
  {
    id: 'plantuml_c4',
    label: 'PlantUML C4 Kiến trúc',
    engine: 'plantuml',
    code: `@startuml
!include <C4/C4_Context>
title Kiến trúc Office Studio Local-Only

Person(user, "Nhân sự Văn phòng", "Soạn thảo & xử lý chứng từ nội bộ")
System(studio, "Office Studio 127.0.0.1", "Bộ công cụ Word, Excel, PDF, Slide chạy thuần máy cá nhân")
System_Ext(printer, "Máy in Mạng / USB", "In ấn trực tiếp chuẩn khổ A4")

Rel(user, studio, "Thao tác trên trình duyệt nội bộ")
Rel(studio, printer, "Lệnh in im lặng chuẩn A4")
@enduml`,
  },
  {
    id: 'drawio_overview',
    label: 'Draw.io XML Sơ đồ khối',
    engine: 'drawio',
    code: `<mxfile host="Office Studio Local" version="21.0.0">
  <diagram id="arch" name="Workflow Architecture">
    <mxGraphModel dx="1200" dy="700" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1200" pageHeight="700" background="#FFFFFF">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="u1" value="Nhân sự" style="shape=umlActor;verticalLabelPosition=bottom;verticalAlign=top;html=1;outlineConnect=0;fillColor=#FFFFFF;strokeColor=#0F172A;strokeWidth=1.8;" vertex="1" parent="1">
          <mxGeometry x="100" y="210" width="60" height="90" as="geometry"/>
        </mxCell>
        <mxCell id="hub" value="Office Studio Hub (Local)" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#2563EB;strokeColor=#1D4ED8;fontColor=#FFFFFF;fontStyle=1;fontSize=14;" vertex="1" parent="1">
          <mxGeometry x="290" y="220" width="220" height="70" as="geometry"/>
        </mxCell>
        <mxCell id="out1" value="Báo cáo Word / PDF-A" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#16A34A;strokeWidth=1.8;fontSize=13;" vertex="1" parent="1">
          <mxGeometry x="640" y="220" width="200" height="70" as="geometry"/>
        </mxCell>
        <mxCell id="e1" edge="1" parent="1" source="u1" target="hub" style="edgeStyle=orthogonalEdgeStyle;rounded=1;html=1;strokeWidth=1.8;strokeColor=#475569;exitX=1;exitY=0.5;entryX=0;entryY=0.5;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        <mxCell id="e2" value="Xuất bản 300 DPI" edge="1" parent="1" source="hub" target="out1" style="edgeStyle=orthogonalEdgeStyle;rounded=1;html=1;strokeWidth=1.8;strokeColor=#16A34A;exitX=1;exitY=0.5;entryX=0;entryY=0.5;labelBackgroundColor=#ffffff;">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>`,
  },
  {
    id: 'canvas_flow',
    label: 'Canvas JSON Luồng xử lý',
    engine: 'canvas',
    code: JSON.stringify(
      {
        canvas: { width: 1200, height: 640 },
        nodes: [
          { id: 'user', title: 'Đầu vào Hồ sơ', type: 'actor', x: 80, y: 240, width: 80, height: 90, icon: 'user' },
          { id: 'proc', title: 'Kiểm định OpenXML', type: 'primary', x: 290, y: 245, width: 200, height: 80, icon: 'shield' },
          { id: 'dec', title: 'Hợp lệ QA?', type: 'decision', x: 600, y: 235, width: 140, height: 100 },
          { id: 'done', title: 'Xuất Bản Sạch', type: 'terminator', x: 880, y: 255, width: 170, height: 60, icon: 'zap' },
        ],
        edges: [
          { source: 'user', target: 'proc', label: 'Tải lên Local', exit_port: 'right', entry_port: 'left' },
          { source: 'proc', target: 'dec', label: '13 Cổng QA', exit_port: 'right', entry_port: 'left' },
          { source: 'dec', target: 'done', label: 'Đạt chuẩn', exit_port: 'right', entry_port: 'left' },
        ],
      },
      null,
      2,
    ),
  },
];

export const CODE_TO_DIAGRAM_SAMPLE = `-- Bảng cơ sở dữ liệu mẫu SQL (Chuyển tự động thành ERD PNG / Draw.io / SVG)
CREATE TABLE phong_ban (
  ma_pb VARCHAR(20) PRIMARY KEY,
  ten_pb VARCHAR(120) NOT NULL
);

CREATE TABLE nhan_vien (
  ma_nv VARCHAR(20) PRIMARY KEY,
  ho_ten VARCHAR(150) NOT NULL,
  ma_pb VARCHAR(20) REFERENCES phong_ban(ma_pb),
  email VARCHAR(120)
);

CREATE TABLE bang_luong (
  id SERIAL PRIMARY KEY,
  ma_nv VARCHAR(20) REFERENCES nhan_vien(ma_nv),
  thang VARCHAR(10),
  thuc_linh NUMERIC(14, 0)
);`;
