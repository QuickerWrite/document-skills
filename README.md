# QuickerWrite documents-skills

Private first-party document skills and stateless rendering worker. No application credentials or user documents belong here.

PPT: neutral HTML → Chromium layout → editable PPTX → LibreOffice PDF → PDFium page images → AI review → version-checked element patches. Word: native DOCX → the same rendering and review; never HTML conversion.

Install `pip install .` with LibreOffice Writer/Impress/Calc, Chromium, Noto CJK and Liberation fonts. `qw-documents-worker --format docx < input.docx > output.pdf` runs without database access. PDFium render helpers and PPT element inspection/patching are reusable by the queue adapter.

The server's dedicated `worker-office` queue owns authentication, cancellation, storage and revision checks. Its release vendors an exact Git commit of this package, verified by a SHA-256 source manifest. qwctl installs OS dependencies, validates Python/rendering dependencies and manages worker-office as a separate service/container. Production does not need GitHub credentials. Updating the lock requires a new server bundle and normal qwctl release installation.

Dashi and Guizang remain separate AGPL engine services; their implementations are not included here. HTML conversion remains in the server adapter in this initial split. Skills are versioned here and synchronized into API builds.
