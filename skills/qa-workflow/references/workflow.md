# Workflow and conversation contract

## Step 1 — Intake prompt

If sources are not provided yet, ask:

> Bạn vui lòng gửi URD, BRD hoặc Confluence Page cần phân tích. Nếu có, hãy
> gửi thêm User Story, Business Rules, API Specification và cho biết chức
> năng/phạm vi cần kiểm thử.

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
