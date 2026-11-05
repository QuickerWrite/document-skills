# PowerPoint workflows

```sh
document-skills create deck.json deck.pptx
document-skills inspect deck.pptx
document-skills edit deck.pptx revised.pptx --changes changes.json
document-skills render revised.pptx revised.pdf
document-skills preview revised.pdf slide.png --page 1
```

Creation payload:
```json
{"title":"Project overview","content":[{"kind":"slide","text":"Progress","meta":{"bullets":["Completed work","Next steps"]}}],"styles":{"body":{"font_name":"Noto Sans","font_size":22}}}
```

Use `render_ppt(title, content, styles, ppt_options=None, template_bytes=None)` in Python for templates and embedded image bytes. Templates retain master/layout/theme; sample slides are removed during generation. This is generation, not a preservation edit.

Patch payload (substitute an ID returned by inspection):
```json
[{"slide_number":1,"element_id":2,"changes":{"text":"Verified progress","font_size":28,"color":"#17365D"}}]
```

Allowed changes: `x,y,w,h` percentages, `text`, `font_size` (6–144 pt), `font_name`, `color` (#RRGGBB), `bold`, `italic`. Up to 50 changes per operation. Geometry must remain within the slide. Patcher supports top-level elements with explicit geometry and existing text runs; it does not edit groups, charts or inherited placeholder geometry. Preserve the original and compare inspected objects after edits.
