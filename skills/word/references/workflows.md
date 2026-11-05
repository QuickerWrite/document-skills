# Word workflows

```sh
document-skills create report.json report.docx
document-skills inspect report.docx
document-skills edit report.docx revised.docx --changes changes.json
document-skills render revised.docx revised.pdf
document-skills preview revised.pdf page.png --page 1
```

Minimal creation payload:
```json
{"title":"Project report","content":[{"kind":"heading","text":"Results","meta":{"level":1}},{"kind":"paragraph","text":"A concise account of verified results."},{"kind":"list","items":["First finding","Next action"]}],"styles":{"body":{"font_name":"Noto Sans","font_size":11}}}
```

Other native generation kinds include table, image (embedded bytes through Python API), quote, code, page_break, horizontal_rule and spacer. Advanced layout options are implemented by `render_word(title, content, styles, options=None)`; consult the source for detailed existing generation schema rather than guessing property names.

Edit payload (body paragraph and run indices start at zero):
```json
[{"paragraph":1,"run":0,"text":"Updated results"}]
```

Inspect first: the document title also occupies a paragraph. Tables appear separately in inspection and are not edited by this run-edit interface. Headers, footers, comments, tracked changes and field evaluation are not covered by the inspector/editor. PDF rendering requires LibreOffice and matching fonts.
