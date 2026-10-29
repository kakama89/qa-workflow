<img src="assets/quality-assurance.png" alt="QA Workflow" width="96" align="right">

# QA Workflow — Kiro Power

Kiro Power cho tester trên **Windows**: phân tích URD/BRD/Confluence, review
Acceptance Criteria, viết và review Test Case, rồi xuất file Excel theo
**template riêng của từng tester**.

## Cài đặt (Kiro IDE)

1. Cài [Python 3](https://www.python.org/downloads/) — khi cài, tick
   **Add python.exe to PATH**.
2. Trong Kiro: mở panel **Powers** → **Add Custom Power** →
   **Import power from GitHub** → dán URL repo này → **Install**.
3. Lần chạy đầu, Kiro sẽ hỏi để cài thư viện `openpyxl` — chọn đồng ý.

4. Kết nối Confluence (Cloud): khi power cần đọc trang Confluence lần đầu,
   Kiro mở trình duyệt để bạn đăng nhập Atlassian và **Approve** quyền truy
   cập. Nếu không thấy trình duyệt mở, vào panel **MCP Servers** của Kiro,
   chọn server `atlassian` của power này và kết nối lại.

Cập nhật bản mới: mở power trong panel **Powers** → **Check for updates** →
**Install updates**.

### Confluence

- Dùng MCP server chính thức của Atlassian (`https://mcp.atlassian.com/v2/mcp`,
  HTTP + OAuth). Không lưu token trong power; mỗi tester tự approve bằng tài
  khoản của mình và chỉ đọc được những trang mình có quyền.
- Power chỉ bật các tool **đọc** (trang, tìm kiếm, tệp đính kèm, comment,
  Jira issue); mọi tool tạo/sửa/xoá trên Confluence, Jira… đều bị tắt.
- Chỉ hỗ trợ Confluence **Cloud** (`xxx.atlassian.net`). Admin Atlassian của
  công ty có thể phải bật Rovo MCP Server; việc dùng có thể tính Rovo credits.
- Cách dùng: dán link trang Confluence vào chat, ví dụ *"Phân tích trang
  này và viết test case: https://congty.atlassian.net/wiki/spaces/ABC/pages/123456/..."*.

## Sử dụng

1. Thả template Excel (`.xlsx`) của bạn vào một trong các chỗ sau (không có
   template thì power dùng template mặc định):
   - `.kiro\qa-workflow\inbox` trong project đang mở — kéo thả ngay trong
     Explorer của Kiro;
   - `%USERPROFILE%\.kiro\qa-workflow\inbox` — dùng cho mọi project;
   - hoặc đính kèm khi chat.
2. Trong chat Kiro, gõ `/qa-workflow` hoặc nói tự nhiên, ví dụ:
   *"Phân tích BRD này và viết test case"*.
3. Làm theo 5 bước: gửi tài liệu → xem phân tích → review Test Case →
   duyệt → nhận file Excel.

Đổi template: thả file mới vào thư mục trên, lần chạy sau tự áp dụng. Mỗi
tester có template riêng trên máy mình; template đặt trong
`.kiro\qa-workflow\inbox` của project sẽ dùng chung cho cả team khi project
được commit lên git.

## Chia sẻ cho team

Đẩy repo này lên GitHub (public, hoặc private nếu mọi người đều có quyền
truy cập), rồi gửi URL cho tester khác để họ import như bước Cài đặt.

> Không đưa template có token, cookie, mật khẩu thật vào repo. Power sẽ từ
> chối template chứa thông tin dạng này.

## Cấu trúc

```text
plugin.json                       manifest của Kiro Power
POWER.md                          metadata hiển thị trong panel Powers (tên, icon)
assets/quality-assurance.png      icon của power (PNG 512x512)
mcp.json                          kết nối Atlassian MCP (Confluence/Jira, chỉ đọc)
skills/qa-workflow/
├── SKILL.md                      hướng dẫn workflow cho agent
├── references/                   quy trình, checklist, quy tắc template
├── scripts/                      đọc template, xuất và kiểm tra Excel (Python)
└── assets/template.xlsx          template mặc định
```
