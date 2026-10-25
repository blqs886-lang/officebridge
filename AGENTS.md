# OfficeBridge development

Build an offline-first, client-side compatibility inspector for PPTX and XLSX.
Preserve editable content. Never rasterize slides, replace formulas with values,
strip unknown package parts, or claim Windows/Office validation without evidence.

Original uploads are immutable. Only explicit repair choices may change a copy.
Every modified XML part and attribute must be recorded. Block signed, encrypted,
malformed, unsupported, or unsafe packages from repair. Read-only analysis may be
allowed when it can be done safely. Network requests must never contain documents.

Keep dependencies small, tests meaningful, examples synthetic, and reports clear
about heuristics. No bundled proprietary fonts. MIT license for our code.

Temporary material belongs in work/ (gitignored). No user files in the repository.
