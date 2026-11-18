Run order (Python 3, pandoc, python-docx, lxml, matplotlib, pillow):
1. python prep.py       -> ms2.md, media/, refs_raw.json, secs.json, tables.json
2. python fig2.py       -> (optional) model schematic figure; not used in the final version
3. python build.py      -> Manuscript_AI_for_Science_restructured.docx
Files: omml.py (LaTeX->editable Word equation via pandoc), ieee.py (APA->IEEE references),
build.py (assembles the document). build.py expects the folders media/media/ and out/.
