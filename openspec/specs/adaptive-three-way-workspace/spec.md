# adaptive-three-way-workspace Specification

## Purpose
TBD - created by archiving change preserve-svn-merge-semantics-and-folded-view. Update Purpose after archive.
## Requirements
### Requirement: Mode-specific workspace chrome
The system SHALL distinguish merge scenarios using only the surrounding workspace chrome while preserving spreadsheet canvases and current data-state rendering.

#### Scenario: Two-way comparison uses gray
- **WHEN** the tool is in two-way comparison mode
- **THEN** the surrounding workspace SHALL use the current gray background

#### Scenario: Update conflict uses pink
- **WHEN** the tool is in update-conflict mode
- **THEN** the surrounding workspace SHALL use a soft pink background distinguishable from white spreadsheet canvases and conflict colors

#### Scenario: Cross-branch merge uses green
- **WHEN** the tool is in cross-branch merge mode
- **THEN** the surrounding workspace SHALL use a soft green background distinguishable from white spreadsheet canvases and success/selection colors

#### Scenario: Spreadsheet rendering is unchanged
- **WHEN** any mode-specific chrome color is applied
- **THEN** main spreadsheet cells SHALL retain the current white base, difference, conflict, cursor, selection, and action-state colors

### Requirement: Fold a proven redundant pane
The system SHALL hide a proven redundant three-way pane by default while retaining its model and offering an immediate expansion control.

#### Scenario: Mine and Base are equivalent
- **WHEN** complete package comparison proves Mine and Base equivalent
- **THEN** one redundant pane SHALL be folded and the visible strip SHALL explain the equality

#### Scenario: User expands all panes
- **WHEN** the user activates `展开三方`
- **THEN** Base, Mine, and Theirs SHALL all be visible without recomputation or loss of merge state

#### Scenario: No pair is equivalent
- **WHEN** no complete pairwise equality is proven
- **THEN** the application SHALL retain the full three-pane layout

### Requirement: Preserve workspace state across folding
The system SHALL preserve user context when switching between folded and expanded three-way presentation.

#### Scenario: Expand after selecting a conflict
- **WHEN** the user expands or folds panes after selecting a sheet and conflict cell
- **THEN** the selected sheet, logical cell, scroll position, pending choices, and initialized merged result SHALL remain unchanged

### Requirement: Explain automatic merge outcome
The system SHALL show a single startup outcome dialog after three-way role analysis and automatic processing.

#### Scenario: Whole-workbook convergence
- **WHEN** the result is initialized from Mine, Theirs, or a common workbook
- **THEN** the dialog SHALL identify the equality evidence, selected result, folded pane, and required next user action

#### Scenario: Semantic pre-merge with remaining conflicts
- **WHEN** supported changes were automatically merged and unresolved conflicts remain
- **THEN** the dialog SHALL show automatic and unresolved counts and SHALL offer navigation to the first unresolved conflict

#### Scenario: Fully automatic semantic merge
- **WHEN** all supported differences are merged and no unresolved conflict remains
- **THEN** the dialog SHALL explain that the result is ready for review and SHALL focus the Save Merged action after dismissal

#### Scenario: Cross-branch revision is already present
- **WHEN** all incoming source changes already exist in Target Working
- **THEN** the dialog SHALL state that no workbook data change is required and SHALL not describe target-only branch differences as automatically merged source changes

### Requirement: Scenario-specific identity labels
The system SHALL use role labels that reflect the launch semantics.

#### Scenario: Cross-branch identity headers
- **WHEN** the launch is a cross-branch merge
- **THEN** pane headers and the outcome dialog SHALL use Source Before, Source After, Target Working, and Target Pristine labels with source repository path/revision evidence when available

#### Scenario: Update-conflict identity headers
- **WHEN** the launch is an update conflict
- **THEN** the existing Base, Mine, Theirs, and target-pristine meanings SHALL remain available

### Requirement: Analysis-time read-only gate
The system SHALL prevent merge-changing actions until role discovery, author lookup, equivalence comparison, and automatic processing reach a terminal state.

#### Scenario: User interacts during analysis
- **WHEN** startup merge analysis is incomplete
- **THEN** the user MAY inspect progress but SHALL NOT change panes, sheets, only-difference state, or merged content

#### Scenario: Analysis fails safely
- **WHEN** startup analysis fails or is cancelled
- **THEN** the system SHALL enter full manual three-way mode with no automatic convergence and SHALL explain the fallback

### Requirement: Actionable structural column conflicts
The system SHALL expose safe explicit column choices as soon as an actionable structural column conflict is ready.

#### Scenario: Target Working is missing a source column
- **WHEN** a ready Sheet has a uniquely identified structural column block and Target Working lacks that logical column
- **THEN** the system SHALL automatically select the first unresolved structural block and enable the applicable Target Working, Source Before, and Source After column choices

#### Scenario: Structural mapping is ambiguous
- **WHEN** the selected structural block has a conservative or ambiguous mapping
- **THEN** the system SHALL retain authoritative revalidation and explicit confirmation before writing and SHALL explain why an action cannot proceed instead of silently disabling every choice

### Requirement: Navigate only to real workbook cells
The system SHALL distinguish workbook-level review markers from navigable Sheet cells.

#### Scenario: First unresolved marker is workbook-level
- **WHEN** the first unresolved marker uses a pseudo Sheet such as `<workbook>`
- **THEN** cell navigation SHALL skip it and continue to the first unresolved marker on a real displayed Sheet

#### Scenario: No real conflict cell is navigable
- **WHEN** unresolved work remains but all markers are workbook-level or structural summaries
- **THEN** the outcome dialog SHALL offer full three-way/manual review and SHALL NOT attempt to select a pseudo Sheet

### Requirement: Compact and legible merge workspace
The system SHALL prioritize spreadsheet rows and semantic identity information over duplicated navigation and diagnostic text.

#### Scenario: Long identity paths
- **WHEN** Source Before, Target Working/Merged, or Source After has a long local or repository path
- **THEN** the visible identity strip SHALL keep role, filename/revision, and Author legible while the full path remains available through hover details or diagnostics

#### Scenario: Sheet navigation
- **WHEN** a workbook is displayed
- **THEN** exactly one visible Sheet navigation strip SHALL be present and it SHALL continue to control the hidden Sheet host

#### Scenario: Vertical workspace allocation
- **WHEN** the main window is maximized or resized
- **THEN** redundant top diagnostics and unused lower-panel space SHALL not reserve rows, the hover title and current-cell identity SHALL share one line, and released height SHALL expand the main spreadsheet grid

### Requirement: Action-first structural toolbar
The system SHALL keep structural merge actions visible and shall guide the user through every unresolved structural block.

#### Scenario: Structural summary is long
- **WHEN** one or more whole-column changes produce a verbose difference summary
- **THEN** the Target Working, Source Before, and Source After column buttons SHALL remain fully visible while the summary is compacted and its complete text remains available through secondary detail

#### Scenario: One of multiple structural blocks is completed
- **WHEN** the selected structural block is applied or retained and another actionable structural block remains
- **THEN** the next block SHALL become selected automatically, the column buttons SHALL remain enabled, and the status SHALL state that more structural work remains

#### Scenario: Final structural block is completed
- **WHEN** no actionable structural block remains after a column decision
- **THEN** the column buttons SHALL become disabled and the status SHALL explicitly state that all structural column differences are handled

### Requirement: Proximate merge guidance and centered navigation
The system SHALL visually separate difference navigation from structural-column decisions while keeping each decision's guidance adjacent to its controls.

#### Scenario: Difference navigation is displayed
- **WHEN** a Sheet comparison view is ready
- **THEN** Previous Difference, the current/total and pending counts, and Next Difference SHALL form one centered control group independent of surrounding summary lengths

#### Scenario: A structural column is pending
- **WHEN** a structural block such as Excel-style logical column T is selected automatically or explicitly
- **THEN** its compact pending status SHALL appear immediately to the left of the three column-action buttons and the logical-column token SHALL be rendered in red

### Requirement: Current-Sheet global cell adoption
The system SHALL offer a Global Mode in both side-action menus for applying all safely mapped cell differences on the current Sheet.

#### Scenario: User chooses Global Mode
- **WHEN** the current Sheet is READY and the user invokes a left-side or right-side action in Global Mode
- **THEN** the tool SHALL operate on the authoritative full-Sheet difference set rather than only visible rows or the current difference block

#### Scenario: Global action is confirmed
- **WHEN** all target cells are safely mapped and no unresolved structural prerequisite blocks the chosen direction
- **THEN** all applicable cell values on the current Sheet SHALL be applied atomically and recorded as one undo operation

#### Scenario: Global action cannot be represented safely
- **WHEN** any target location is stale, unresolved, structurally ambiguous, or otherwise unsafe
- **THEN** the tool SHALL write nothing and SHALL explain the blocker rather than partially applying the Sheet

### Requirement: Stable per-Sheet uniform column geometry
The system SHALL render every logical column in one Sheet with the same bounded display width across the complete view.

#### Scenario: Cell content lengths differ by row or side
- **WHEN** rows or Base/Mine/Theirs panes contain different-length values in the same logical column
- **THEN** every logical column in that Sheet SHALL use one common bounded width across every main row, pane, and A-Z/AA header rather than deriving separate widths from cell contents

#### Scenario: The projection changes
- **WHEN** a row/column action or undo changes the authoritative Sheet projection
- **THEN** the width model SHALL be rebuilt with the projection while remaining stable during ordinary scrolling, only-difference filtering, and cursor movement

### Requirement: Excel-style comparison column labels
The system SHALL use Excel-style logical column labels consistently in user-facing grid and hover-comparison headers.

#### Scenario: Hover comparison renders logical columns
- **WHEN** C-area comparison columns are displayed
- **THEN** their labels SHALL use A through Z, then AA and later Excel-style names matching the main view rather than internal L1, L2, or L100 identifiers

### Requirement: Pixel-stable mixed-language column geometry
The system SHALL preserve one pixel-aligned logical-column boundary across rows containing Latin, numeric, Chinese, and empty values.

#### Scenario: Latin and Chinese rows share a column
- **WHEN** one fixed-width Sheet contains rows such as `id@id`, `uint32`, `小等级`, `此行不填`, and `1`
- **THEN** every rendered column boundary SHALL align with the same Excel-style header boundary across every row and visible pane

#### Scenario: Selection and difference styling changes
- **WHEN** a row or cell receives cursor, selection, conflict, or difference styling
- **THEN** the styling SHALL NOT substitute a font whose metrics move column boundaries

### Requirement: One compact merge-action row
The system SHALL place difference navigation and structural-column decisions in one vertically compact action row.

#### Scenario: Normal-width workspace
- **WHEN** the Sheet toolbar has sufficient horizontal space
- **THEN** Previous Difference/status/Next Difference and the structural status plus three column buttons SHALL share one row without overlap or redundant blank rows

#### Scenario: Narrow workspace
- **WHEN** the combined controls cannot remain at their preferred positions
- **THEN** the navigation group SHALL shift within the available left region while structural actions remain visible and no controls overlap or clip

### Requirement: Left-aligned root utility actions
The system SHALL place the root-level utility buttons at the far left in a stable order.

#### Scenario: Root toolbar is displayed
- **WHEN** the application workspace is created
- **THEN** `重算并刷新`, `导出诊断包`, `复制反馈信息`, and `检查更新` SHALL appear from left to right at the left edge without a leading expanding spacer

### Requirement: Excel labels in all user-facing column guidance
The system SHALL reserve `L<n>` notation for internal diagnostics and use Excel-style labels in visible structural-column guidance.

#### Scenario: Structural column is pending or completed
- **WHEN** a status, tooltip, confirmation, blocker, conflict location, or completion message references a logical column
- **THEN** the user-facing text SHALL show `A..Z`, `AA..ZZ`, and later Excel-style labels and SHALL NOT show `L1..L100`
### Requirement: Explicit cache-backed difference browser

The system SHALL expose a default-expanded difference browser built from the existing per-Sheet comparison cache, with no second workbook scan.

#### Scenario: browse and locate a difference

- **WHEN** the user filters or selects a difference row
- **THEN** the browser shows type/count/status and synchronizes the selected Sheet and real workbook cell

#### Scenario: double-click safety

- **WHEN** the user double-clicks a workbook cell or difference row
- **THEN** the tool only selects/locates and displays details; it SHALL NOT write workbook content

### Requirement: Unsaved close protection

The system SHALL guard window close when either editable side is dirty.

#### Scenario: save, discard or cancel

- **WHEN** the user closes with unsaved changes
- **THEN** the tool offers “保存并关闭 / 放弃改动并关闭 / 取消”; a failed save leaves the window and edits intact

### Requirement: Consistent two-way roles

The system SHALL use one role presentation for identity cards, pane titles and direction-labelled actions.

#### Scenario: two-way comparison

- **WHEN** comparing two files without a Base workbook
- **THEN** the left side is labeled Base and the right side Mine consistently, with no contradictory Theirs label
