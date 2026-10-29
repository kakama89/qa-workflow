# Workflow and conversation contract

## Mode selection

When the request does not make the mode clear, ask:

> Bạn muốn làm theo quy trình nào?
>
> 1. **Quy trình bình thường** — phân tích tài liệu yêu cầu (URD, BRD,
>    Confluence, User Story) rồi viết Test Case theo yêu cầu.
> 2. **Characterization Test** — chụp lại hành vi hiện tại của hệ thống đang
>    chạy làm mốc hồi quy (trước khi refactor, migrate, nâng cấp, hoặc khi
>    tài liệu thiếu/lỗi thời).
> 3. **Kết hợp cả hai** — phần có tài liệu viết theo yêu cầu, phần chưa có
>    tài liệu chụp hành vi hiện tại; tách thành các phần riêng.
>
> Bạn trả lời số 1, 2 hoặc 3. Chưa chắc thì cho mình biết bạn đang có tài
> liệu yêu cầu hay chỉ có hệ thống đang chạy, mình sẽ gợi ý.

When the mode is clear from the request, confirm in one line instead, for
example:

> Mình làm theo **Characterization Test** (chụp hành vi hiện tại). Muốn đổi
> sang quy trình bình thường thì trả lời "1".

For option 3, run Step 2 as both a Requirement Analysis Report (for parts
with documents) and a Behavior Inventory (for the rest), and keep the two
kinds of test cases in separate sections through Steps 3–5.

## Step 1 — Intake prompt

If sources are not provided yet, ask:

> Bạn vui lòng gửi URD, BRD hoặc Confluence Page cần phân tích. Nếu có, hãy
> gửi thêm User Story, Business Rules, API Specification và cho biết chức
> năng/phạm vi cần kiểm thử.

For a Characterization Test (baseline of current behavior), ask instead:

> Bạn muốn chụp lại hành vi hiện tại của chức năng nào (màn hình, API, batch)?
> Cho mình biết build/version và môi trường, lý do (refactor, migrate…), và
> gửi bằng chứng đang có: screenshot, request/response API, log, kết quả truy
> vấn DB hoặc file xuất ra.

Settle the template without asking when possible: run `qa_profiles.py auto`
first, then state the result in one line, for example:

> Mình dùng template `template-ngan.xlsx` (sheet Testcases: ID=A, Title=B,
> Steps=D, Expected=F, Priority=H). Nếu muốn dùng template khác, bạn chỉ cần
> thả file vào `.kiro\qa-workflow\inbox` trong project (hoặc
> `%USERPROFILE%\.kiro\qa-workflow\inbox` để dùng cho mọi project), hoặc gửi
> kèm trong chat.

Ask only when needed:

- Several templates, no clear choice: "Mình thấy các template: `<danh sách>`.
  Bạn dùng template nào?"
- No template anywhere: say the bundled default template is used and where
  to drop their own next time.
- Template refused (possible credentials, value not shown) or columns not
  recognized: explain briefly and ask for a fixed file or the column names.

Identify source/version/scope. Ask only for details that materially affect
the analysis. After profiling the template, show the detected layout in a
short table (sheet, header row, column → role, ID convention, dropdowns) and
ask the user to confirm it.

## Step 2 — Requirement analysis output

Use this structure:

1. Source inventory
2. Scope and exclusions
3. Requirement / AC inventory
4. Findings table: ID, type, severity, source/location, evidence, impact,
   question/recommendation, status
5. AC quality assessment
6. Risks and assumptions
7. Decision questions

Severity guidance:

- **Critical**: security, data integrity, or core flow can be materially
  compromised.
- **High**: major flow or expected behavior cannot be determined.
- **Medium**: important edge case, validation, or exception is unspecified.
- **Low**: wording, documentation, or minor testability issue.

Severity is a suggested assessment, not a substitute for business
prioritization.

In Characterization Test mode, output a Behavior Inventory instead (see
`characterization.md`): source/evidence inventory, scope, observed behaviors
with evidence IDs and stability, masked non-deterministic fields, suspected
defects, gaps, and decision questions.

## Step 3 — Test Case Review output

Present a compact table, split into several tables if needed:

| ID | Requirement/AC | Title | Type | Preconditions | Test Data | Steps | Expected Result | Priority | Issue/Assumption | Review Status |
|----|----------------|-------|------|---------------|-----------|-------|-----------------|----------|------------------|---------------|

- Use `TBD` only when the value is known to be required; explain what is
  missing.
- Use `N/A` only when genuinely not applicable.
- Keep a traceable link from each test case to its source requirement where
  possible.
- Use the template's own dropdown values (for example Priority
  `High/Medium/Low`) so the export does not produce invalid values.

## Step 4 — Approval prompt

Ask the user to choose:

- Approve Test Cases and proposed template changes
- Request changes to Test Cases
- Request changes to template
- Return to requirement clarification

In Characterization Test mode, also ask for a decision on each
`Suspected defect — confirm` and `Varies` behavior: keep as baseline, defect
(baseline as-is, log bug), or exclude.

If the user requests changes, apply them and show the revised draft.
Approval applies only to the version shown; material changes afterward
require another review.

## Step 5 — Export handoff

Deliver:

- Path (or download link, where the platform provides one) of the new
  `.xlsx` file
- Sheets added/removed/renamed
- Columns or dropdowns changed
- Formula/coverage changes
- Validation results (errors fixed, warnings remaining, checks that could not
  run)
- Remaining unresolved issues and assumptions

Do not hide unresolved issues. If the user approves export with open issues,
flag them in the workbook.
