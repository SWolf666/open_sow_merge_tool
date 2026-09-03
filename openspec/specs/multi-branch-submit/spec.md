# Multi-branch submit

## Scenario: select branches and files

- **WHEN** the user opens branch submit mode
- **THEN** the tool discovers the SVN working-copy root, dynamically shows all same-repository top-level branches including `master` in descending SVN change-time order, and requires one source, one or more targets, and one or more safe `.xlsx` changes

## Scenario: recursive multi-file workbench

- **WHEN** the user starts from an `.xlsx`, folder, or folder background
- **THEN** the tool recursively scans that folder, shows all SVN changes with status metadata, defaults safe modified/added/deleted Excel files on, keeps unversioned Excel off, and displays non-Excel changes read-only

## Scenario: mandatory source-change preflight

- **WHEN** one or more target branches are selected
- **THEN** the user must run “预检查（必需）”; the tool obtains source pristine and working content, computes source-before to source-after changes, and previews a per-target action matrix without changing target files

## Scenario: native single-branch submission remains unchanged

- **WHEN** the user uses TortoiseSVN's original single-branch commit entry without selecting multi-branch targets
- **THEN** the original TortoiseSVN flow remains available and is not gated by this workbench's preflight

## Scenario: fail closed

- **WHEN** status collection fails, a selected target is dirty, conflict/switch/property/external/unsupported structure is detected, or a target changed since analysis
- **THEN** the affected action is blocked and no commit dialog is opened for that unsafe action

## Scenario: conservative fast path

- **WHEN** a source delta is limited to value/formula changes or unique-key rows appended at the source tail, with stable headers
- **THEN** the tool classifies the target in read-only OOXML, projects only the proven cells/rows into a target-derived candidate using target-side styles, preserves unrelated target content, and reports direct/already/confirmation counts

## Scenario: inspect actual target modification points

- **WHEN** the user selects a mappable modified file in the preflight matrix and chooses “查看目标修改点”
- **THEN** the tool lazily creates an artifact-only preview derived from the unchanged target, opens the comparer as target-before versus target-after-preview, never compares the source workbook directly as the target result, and does not modify the target working copy

## Scenario: overlapping target content requires confirmation

- **WHEN** a target record or field differs from both source-before and source-after while the source-change location remains technically mappable
- **THEN** the tool marks the file “需人工确认”, shows source-before/source-after/target values, and requires the user either to accept all source changes for that target/file pair or explicitly exclude it from the batch

## Scenario: unsupported workbook changes remain blocked

- **WHEN** rows are inserted away from the proven tail, protected workbook structures change, or a unique record/field mapping cannot be proved
- **THEN** the tool marks the action “安全阻断”; confirmation alone cannot convert an unsupported write into a successful synchronization

## Scenario: ignore save-time presentation churn

- **WHEN** Excel or WPS rewrites selection/view metadata, column widths, comment relationships, conditional-format identifiers, producer extensions, or equivalent CRLF encodings without changing configuration cell semantics
- **THEN** the fast analyzer ignores that churn while still blocking real merge-cell, data-validation, table, protection, embedded-object, VBA, or defined-name changes

## Scenario: added, deleted, and renamed files

- **WHEN** an added or deleted Excel file is selected
- **THEN** add requires an absent target path and versioned parent; a different existing target file is blocked because there is no common baseline; delete is direct when the target equals the deletion baseline and otherwise requires explicit delete confirmation; a detected rename is blocked with a Repair Move instruction

## Scenario: source commit first

- **WHEN** the user starts a ready batch
- **THEN** the source branch commit dialog opens first with the frozen user message, and targets remain pending until every file is reconciled against SVN status, pristine, hashes, and revisions after the dialog closes

## Scenario: partial source selection

- **WHEN** the user commits only some selected source files
- **THEN** propagation stops, the committed files are split into an explicit child batch, and the remaining source files return to the workbench

## Scenario: per-target submission

- **WHEN** a target is ready after revalidation
- **THEN** the tool writes a durable prepare intent and backup, applies only the planned candidate, opens one TortoiseSVN commit dialog with the frozen message and tracking footer, and marks each file committed only after post-commit reconciliation

## Scenario: modified targets are never replaced by source files

- **WHEN** a modified-file action is prepared
- **THEN** the candidate must exist, be derived from the exact target snapshot, and contain only accepted source-change locations; a missing or changed candidate blocks the action instead of falling back to copying the complete source workbook

## Scenario: cancellation and recovery

- **WHEN** any source or target commit is cancelled or fails
- **THEN** the tool stops subsequent targets, persists per-file facts, keeps completed files immutable, and on resume reconciles actual SVN state before opening another dialog or restoring only hash-matching uncommitted candidates

## Scenario: legacy compatibility

- **WHEN** the tool is launched with existing two-way, three-way, or TortoiseSVN diff/merge arguments
- **THEN** those paths retain their current behavior and do not enter branch submit mode

## Scenario: Explorer context menu

- **WHEN** the current user invokes “多分支 SVN 提交” from an `.xlsx`, folder, or folder background
- **THEN** the tool opens branch-submit mode, infers the working-copy/source/scope, and recursively scans the selected folder without changing legacy TortoiseSVN registrations

## Scenario: context-menu uninstall

- **WHEN** the current user runs the dedicated context-menu uninstaller
- **THEN** only the three owned `SowMultiBranchSVNSubmit` shell trees are removed and other Explorer or TortoiseSVN settings remain untouched

## Scenario: target-free mode is native commit

- **WHEN** no target branch is selected
- **THEN** the workbench shows “源分支单提交” and “SVN 单分支提交”, skips multi-branch preflight and opens the native source-branch commit flow

## Scenario: busy request inputs are frozen

- **WHEN** scanning, preflight, recovery or commit is active
- **THEN** source, target selection, scope and file selection cannot change until the operation completes or is cancelled

## Scenario: keyboard-first workbench

- **WHEN** the workbench has focus
- **THEN** Space toggles a selected checkbox, Enter invokes the enabled primary action, Escape cancels/returns, and Ctrl+F focuses branch search
