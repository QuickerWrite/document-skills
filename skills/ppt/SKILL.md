---
name: ppt
description: Generate editable PowerPoint slides, inspect object identities, patch selected objects while preserving other package parts and verify rendered slides.
---
# PowerPoint

Use the standalone `document-skills` CLI or `document_skills.render_ppt` / `document_skills.pptx_patch`. Read [workflows](references/workflows.md) for creation and patch examples.

Inspect the presentation before edits. Use the returned slide number and element ID; IDs must never be guessed from text or visual order. Geometry is percent of slide dimensions, font sizes are points. Patch only requested objects. Unchanged ZIP members are preserved by the patcher. Text replacement adopts the first existing run and clears later run text; it does not preserve mixed formatting within the replacement. Inherited geometry, groups and empty text elements can require a dedicated editor.

Prefer native editable text, tables and shapes. Reuse the user's theme, maintain hierarchy and keep body text readable at presentation distance. Split overcrowded slides instead of shrinking everything. Do not claim a screenshot is an editable slide.

Reopen and inspect generated output, render all slides to PDF, inspect PNG previews and correct clipping, overlap, missing glyphs or inconsistent alignment. Structural checks alone are insufficient. Report unsupported objects, unavailable rendering or target-viewer differences explicitly.
