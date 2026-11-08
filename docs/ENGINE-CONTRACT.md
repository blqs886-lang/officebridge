# Engine integration contract

`analyzeOffice(bytes: Uint8Array, fileName: string, options = {})` -> report
`repairOffice(bytes, fileName, { selectedIssueIds: string[], ...options })`
-> `{ bytes: Uint8Array, fileName: string, report, changes: Change[] }`

Report fields: `schemaVersion: 1`, `toolVersion`, `fileName`, `format: 'pptx'|'xlsx'`,
`sha256`, `fileSize`, `issues: Issue[]`, `inventory`, `repairBlocked: boolean`,
`validation: { packageChecked, windowsOfficeTested: false, notes: string[] }`.

Issue: `{ id, code, severity: 'warning'|'info'|'error', title, description,
location: string, part: string, fixable: boolean, defaultSelected: boolean,
fix?: { kind: 'font-replace', from: string, to: string, scope: string, affectsDefaultFont?: true, affectsWrappedRows?: true },
evidence?: string, recommendation?: string }`.

Change: `{ issueId, part, kind, before, after, count, printReview?: string }`. XLSX changes that touch the default font or text in wrapped, automatically sized rows carry their print-preview guidance into the post-repair report.

Format scanner exports `scanPptx(context)` / `scanXlsx(context)` returning
`{ issues, inventory }`. Context is `{ entries: Map<string, Uint8Array>,
xml: Map<string, Document>, texts: Map<string, string>, options }`.
Options: `{ fontProfile: 'windows-office', fontMappings?: object, targetFonts?: string[] }`. With a target list, reports include only the count and document family matches under `inventory.targetFontComparison`. The caller keeps the list for repair.
Scanners only inspect. Use predictable IDs from code + part + index; avoid random IDs.
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
