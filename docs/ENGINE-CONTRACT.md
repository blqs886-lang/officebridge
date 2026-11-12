# Engine integration contract

`analyzeOffice(bytes: Uint8Array, fileName: string, options = {})` -> report
`repairOffice(bytes, fileName, { selectedIssueIds: string[], ...options })`
-> `{ bytes: Uint8Array, fileName: string, report, changes: Change[] }`

Report fields: `schemaVersion: 1`, `toolVersion`, `fileName`, `format: 'pptx'|'xlsx'`,
`sha256`, `fileSize`, `issues: Issue[]`, `inventory`, `repairBlocked: boolean`,
`validation: { packageChecked, targetPlatform, targetOfficeTested: false, windowsOfficeTested: false, notes: string[] }`.

Issue: `{ id, code, severity: 'warning'|'info'|'error', title, description,
location: string, part: string, fixable: boolean, defaultSelected: boolean,
fix?: { kind: 'font-replace', from: string, to: string, scope: string, layoutScope?: string,
affectsDefaultFont?: true, affectsWrappedRows?: true, affectsCellFit?: true, layoutUnverified?: true },
layoutReview?: string, evidence?: string, recommendation?: string }`.

Change: `{ issueId, part, kind, before, after, count, location?: string, printReview?: string }`.
A split choice records its human-readable location in the change log. An XLSX
selection that touches a recognized layout risk carries its print-preview
guidance into the post-repair report. A user can compare locally exported proofs
to check the actual page count and appearance.

Format scanner exports `scanPptx(context)` / `scanXlsx(context)` returning
`{ issues, inventory }`. Context is `{ entries: Map<string, Uint8Array>,
xml: Map<string, Document>, texts: Map<string, string>, options }`.
Options: `{ fontProfile: 'windows-office'|'mac-office', fontMappings?: object, targetFonts?: string[] }`. With a target list, reports include only the count and document family matches under `inventory.targetFontComparison`. The caller keeps the list for repair.
Scanners only inspect. Issue IDs use the code, part and family ordinal. When a
single source font in one namespace and part has separately selectable groups,
all its choices retain that ordinal and append their own `fix.layoutScope`.
The old whole-family ID no longer selects a subset. Reanalysis during repair
fills a private `context.fontTargets` map from those scoped IDs to exact DOM
elements. A shared declaration belongs to one group, including any layout risks
in all its known uses. Callers select IDs from the current analysis. A missing or
overlapping target stops the repair.
All UI copy in Chinese. English code identifiers. No numeric compatibility score.

Shared `src/engine/xml.js` exports `NS`, `elements(node, ns, localName)`,
`parseXml(text, part)`, `attribute(node, name, fallback = '')`.
NS keys: a, p, s, r, rel, ct. `elements` returns array of matching descendants.
Shared `fonts.js` exports `fontAdvice(name, options)` ->
`null | { to, reason, confidence: 'high'|'medium', defaultSelected: boolean }`.
`fontAdvice` returns null for theme references and known supported/unknown fonts;
unknown fonts must still be surfaced by scanner as unverified (not automatically safe).

Each scanner owns font issues for its package parts. PPTX scans ppt/** and all
DrawingML fonts. XLSX scans xl/**, including drawing/chart/theme fonts.
Font issue fix.scope is an XML namespace URI, font names are `typeface` in DrawingML
and `val` on SpreadsheetML `name` / `rFont`. Scope is used by repair to limit mutation.
Font replacements must be opt-in; defaultSelected false. Substitution may change layout.
Do not rewrite formulas, external paths, media, theme refs, dimensions or unknown elements.

Core rejects macros for repair (pptm/xlsm unsupported), signatures, dangerous ZIP
paths, zip bombs, XML DTD/entities, duplicate entries and malformed relationships.
Existing unknown part bytes are retained. Package-level inventory from core is merged.

## Local Office records

The native helpers produce a separate `officebridge-native-check`, schema version 1 record. It is never merged into the rule scanner's `targetOfficeTested` flags. An imported native record identifies one execution environment and the original/copy fingerprints. The reader verifies embedded PDF data against its SHA-256 before rendering; this is an integrity check, not proof of who authored the record. The result HTML embeds JSON as non-executable, escaped text. Native checks use the installed Office application and are not an untrusted-document sandbox.
