---
name: word
description: Create native Word documents, inspect paragraphs and tables, edit individual text runs and verify rendered layouts with the standalone document-skills toolkit.
---
# Word

Use `document-skills` or the `document_skills.render_word` and `document_skills.word` APIs. Read [workflows](references/workflows.md) for commands and input structure. Run `document-skills doctor` to check the optional renderer.

Inspect before changing an existing document. Run indices are zero-based and refer to body paragraphs; they are not page numbers. Edit individual runs to retain surrounding formatting. Do not rebuild an entire paragraph to replace a few words. The conservative editor rejects runs containing fields or drawings; tracked changes, comments and complex field editing require a dedicated OOXML editor and are not implemented here.

Use semantic headings and actual lists/tables, paragraph spacing rather than blank lines, and the existing template conventions. Do not promise page counts from extracted text. Word page numbers, TOCs and pagination depend on a layout engine and fonts.

Create or edit to a new file, reopen and inspect the result, then render to PDF and examine pages for clipping, broken tables, missing glyphs and poor page breaks. Report renderer/font limitations if visual verification cannot run. Never fabricate citations, signatures or missing facts.
