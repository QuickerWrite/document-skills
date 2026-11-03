"""Stateless Office rendering, shared by the standalone worker and queue adapter."""
import os, re, shutil, signal, subprocess, tempfile, time, zipfile
from pathlib import Path
from io import BytesIO
from defusedxml import ElementTree
MAX_BYTES = 80 * 1024 * 1024
MAX_PAGES = 200
class RenderCancelled(Exception):
    pass
def _bounded_bytes(data: bytes) -> bytes:
    if not data or len(data) > MAX_BYTES:
        raise ValueError("文档为空或超过 80 MB 处理限制")
    return data



def _guard_ooxml(data: bytes) -> None:
    """在 LO 打开之前拒绝解压炸弹、宏和链接资源。"""
    _bounded_bytes(data)
    try:
        with zipfile.ZipFile(BytesIO(data)) as package:
            parts = package.infolist()
            if len(parts) > 12000 or sum(item.file_size for item in parts) > 250 * 1024 * 1024:
                raise ValueError("Office 文档解压后过大")
            for part in parts:
                name = part.filename.lower()
                if part.file_size > 50 * 1024 * 1024:
                    raise ValueError("Office 文档包含超大内部资源")
                if "vbaproject" in name or name.startswith("xl/externallinks/"):
                    raise ValueError("不处理包含宏或外部工作簿链接的 Office 文件")
                if name.endswith(".rels"):
                    root = ElementTree.fromstring(package.read(part))
                    for rel in root:
                        if rel.get("TargetMode") == "External" and not rel.get("Type", "").endswith("/hyperlink"):
                            raise ValueError("请先将文档外部引用资源嵌入文件后再转换")
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    root = ElementTree.fromstring(package.read(part))
                    for node in root.iter():
                        if node.tag.rsplit("}", 1)[-1] == "f":
                            _check_formula("=" + (node.text or ""))
    except zipfile.BadZipFile as exc:
        raise ValueError("Office 文件不是有效的 OOXML 压缩包") from exc


def _check_formula(value: str) -> None:
    if not value.startswith("=") or len(value) > 8192:
        raise ValueError("公式必须以 = 开头，且不能超过 8192 字符")
    if re.search(r"(?:WEBSERVICE|FILTERXML|RTD|DDE|CALL|REGISTER|EXEC|IMPORTDATA|IMPORTXML)\s*\(", value, re.I) or "|" in value or re.search(r"\[[^]]+\].*!", value):
        raise ValueError("不允许访问外部网络、文件或执行程序的公式")


def _libreoffice_pdf(data: bytes, extension: str, cancelled=lambda: False) -> bytes:
    if extension not in {".docx", ".pptx", ".xlsx"}:
        raise ValueError("Unsupported Office format")
    _guard_ooxml(data)
    binary = os.environ.get("LIBREOFFICE_BIN", "").strip() or shutil.which("libreoffice") or shutil.which("soffice")
    if not binary:
        raise RuntimeError("Office PDF 转换引擎未安装，未生成转换结果")
    with tempfile.TemporaryDirectory(prefix="qw-office-convert-") as temporary:
        folder = Path(temporary)
        source = folder / ("source" + extension)
        source.write_bytes(data)
        profile = folder / "profile"
        (profile / "user").mkdir(parents=True)
        (profile / "user" / "registrymodifications.xcu").write_text(
            '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry">'
            '<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>'
            '<item oor:path="/org.openoffice.Office.Common/Load"><prop oor:name="UpdateDocMode" oor:op="fuse"><value>0</value></prop></item>'
            '</oor:items>', encoding="utf-8")
        command = [str(binary), "-env:UserInstallation=" + profile.as_uri(), "--headless", "--nologo", "--nodefault", "--nolockcheck", "--norestore",
                   "--convert-to", "pdf", "--outdir", str(folder), str(source)]
        # Command arguments are fixed and every path points into this task's
        # freshly-created private directory; shell execution is never used.
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)  # nosec B603
        start = time.monotonic()
        try:
            while True:
                try:
                    _, error = process.communicate(timeout=.5)
                    break
                except subprocess.TimeoutExpired:
                    if cancelled():
                        raise RenderCancelled("Office conversion cancelled")
                    if time.monotonic() - start > 180:
                        raise RuntimeError("Office PDF 转换超时，未交付不完整文件")
            target = folder / "source.pdf"
            if process.returncode != 0 or not target.is_file():
                raise RuntimeError("Office PDF 转换失败，请检查文档或转换引擎")
            output = _bounded_bytes(target.read_bytes())
            if not output.startswith(b"%PDF-"):
                raise RuntimeError("转换引擎未返回有效 PDF")
            return output
        finally:
            if process.poll() is None:
                os.killpg(process.pid, signal.SIGKILL)
                process.communicate()


def _pdf_document(data: bytes):
    import pypdfium2 as pdfium
    document = pdfium.PdfDocument(_bounded_bytes(data))
    if not 1 <= len(document) <= MAX_PAGES:
        document.close()
        raise ValueError(f"PDF 必须为 1 至 {MAX_PAGES} 页")
    return document


def _page_png(document, number: int, width: int = 1000) -> tuple[bytes, float, float]:
    page = document[number - 1]
    try:
        page_width, page_height = page.get_size()
        bitmap = page.render(scale=min(width / max(page_width, 1), 1800 / max(page_height, 1)))
        try:
            image = bitmap.to_pil()
            output = BytesIO()
            image.save(output, format="PNG")
            return output.getvalue(), page_width, page_height
        finally:
            bitmap.close()
    finally:
        page.close()


