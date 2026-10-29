# Quality checklist

Items marked *(script)* are checked by `scripts/validate_workbook.py`. Check
the rest manually, and check everything manually if the script cannot run.

## Requirement quality

- [ ] Source and version identified
- [ ] Scope and exclusions stated
- [ ] Contradictions and gaps recorded with source evidence
- [ ] Acceptance Criteria presence and testability assessed
- [ ] Assumptions separated from confirmed facts
- [ ] Open questions assigned a clear decision owner where known

## Test Case quality

- [ ] Each case has a unique ID *(script, for literal or cached IDs)*
- [ ] Requirement/AC traceability exists where possible
- [ ] Steps are clear, atomic, and executable
- [ ] Expected Result is observable and verifiable
- [ ] Relevant positive, negative, boundary, validation, permission,
      exception, and integration cases considered
- [ ] Test Data uses safe placeholders for secrets *(script)*
- [ ] No unsupported business rule was invented
- [ ] No duplicate or redundant cases without justification
- [ ] Review status reflects actual user approval

## Characterization Test (when that mode is used)

- [ ] Build/version, environment, and observation date recorded
- [ ] Every Expected Result is backed by an evidence ID
- [ ] Behavior reproduced, or marked `Varies` / `Not reproduced`
- [ ] Non-deterministic fields (timestamps, generated IDs, ordering) masked
- [ ] Suspected defects labeled and a Behavior Decision recorded for each
- [ ] Gaps (inputs, roles, states without evidence) listed
- [ ] No observation taken from Production with real customer data
- [ ] Test Results left blank or `Untested`

## Template change review

- [ ] Template layout was profiled and confirmed by the user
- [ ] Each sheet/column/formula/dropdown/format change is listed
- [ ] Handling of existing sample rows (clear/keep) is decided
- [ ] Reason and impact are explained
- [ ] User approval recorded before applying material changes
- [ ] Original template remains unchanged *(script)*
- [ ] Existing useful styles, formulas, validations, and named ranges
      preserved where possible

## Export validation

- [ ] File reopens successfully *(script)*
- [ ] Required sheets and headers are present *(script, compared with template)*
- [ ] Required fields (ID, steps, expected result) are filled *(script)*
- [ ] No duplicate Test Case IDs *(script)*
- [ ] ID formulas are consistent across rows *(script)*
- [ ] Summary formulas cover all test case rows *(script)*
- [ ] Dropdowns cover all test case rows and values match the lists *(script)*
- [ ] No `#REF!` in formulas *(script)*
- [ ] Report metrics have documented definitions
- [ ] No fabricated execution results *(script)*
- [ ] No credentials or secrets remain in exported cells *(script)*
- [ ] Traceability links resolve
- [ ] Remaining issues are visible
