"""
LAIM db-connector — упаковывает .docx «Отчёт о разработке» для doc-browser.

Источник отчёта (по приоритету):
  1) порт report_path подключён к Data Source → читаем РЕАЛЬНЫЙ отчёт оттуда.
     Тип файла определяется ПО СИГНАТУРЕ БАЙТОВ, а не по расширению (Data Source/
     getPortAsLocalPath может отдать файл без расширения):
        - ZIP/DOCX (b'PK\\x03\\x04') → берём байты как есть;
        - pickle  (b'\\x80')        → распаковываем: dict с 'docx_bytes' → эти байты;
                                       dict с 'text' / строка / список → собираем .docx из текста.
  2) порт не подключён → последний fallback: собираем минимальный .docx из _DEV_REPORT.

То есть реальный отчёт о разработке (например dev_report_strategybuilder.pkl или .docx)
используется, как только он подан на report_path. _DEV_REPORT — только аварийная заглушка.
"""
import io
import pickle
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import pandas as pd

# ============================================================
# Аварийный fallback-текст (используется ТОЛЬКО если report_path не подключён)
# ============================================================
_DEV_REPORT = [
    "Отчёт о разработке GenAI-решения (fallback-заглушка)",
    "",
    "report_path не подключён — подайте реальный отчёт о разработке на вход report_path",
    "(Data Source с .docx или с .pkl, содержащим ключ 'docx_bytes' или 'text').",
]

# ============================================================
# Минимальный валидный OOXML (.docx) средствами stdlib
# ============================================================
_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _document_xml(paragraphs):
    def _p(text):
        if text == "":
            return "<w:p/>"
        return f'<w:p><w:r><w:t xml:space="preserve">{escape(str(text))}</w:t></w:r></w:p>'
    body = "".join(_p(p) for p in paragraphs)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{_W_NS}"><w:body>{body}<w:sectPr/></w:body></w:document>'
    )


def _build_docx_bytes(paragraphs):
    """Собирает валидный .docx в памяти и возвращает его байты."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("word/document.xml", _document_xml(paragraphs))
    return buf.getvalue()


def _text_to_paragraphs(text):
    return str(text).splitlines() or [str(text)]


def _report_bytes_from_pickle_obj(obj):
    """Pickle-объект отчёта → байты валидного .docx."""
    if isinstance(obj, dict):
        b = obj.get("docx_bytes") or obj.get("bin")
        if b:                                   # уже готовый .docx внутри pickle
            return bytes(b)
        text = obj.get("text") or obj.get("report") or ""
        return _build_docx_bytes(_text_to_paragraphs(text))
    if isinstance(obj, (list, tuple)):
        return _build_docx_bytes([str(x) for x in obj])
    return _build_docx_bytes(_text_to_paragraphs(obj))


def _resolve_to_file(report_path):
    """report_path может быть файлом или директорией Data Source с одним файлом."""
    p = Path(report_path)
    if p.is_dir():
        files = [f for f in p.iterdir() if f.is_file() and not f.name.startswith((".", "~$"))]
        if not files:
            raise FileNotFoundError(f"В директории '{p}' нет файлов.")
        if len(files) > 1:
            raise ValueError(f"В директории '{p}' найдено {len(files)} файлов. Ожидается один.")
        return files[0]
    if p.is_file():
        return p
    raise FileNotFoundError(f"report_path не найден: {p}")


def _load_report_bytes(report_path):
    """Читает реальный отчёт о разработке и возвращает байты валидного .docx.
    Тип определяется по СИГНАТУРЕ: ZIP/DOCX=b'PK\\x03\\x04', pickle=b'\\x80'."""
    f = _resolve_to_file(report_path)
    with open(str(f), "rb") as fh:
        data = fh.read()
    if data[:4] == b"PK\x03\x04":                       # zip/docx
        print(f"FILE EXTRACTION OK (docx): {f.name} ({len(data)} bytes)")
        return data
    if data[:1] == b"\x80" or f.suffix.lower() in (".pkl", ".pickle"):   # pickle
        obj = pickle.loads(data)
        out = _report_bytes_from_pickle_obj(obj)
        print(f"FILE EXTRACTION OK (pickle→docx): {f.name} → {len(out)} bytes")
        return out
    if f.suffix.lower() == ".docx":
        return data
    # последний шанс: pickle, иначе трактуем как текст
    try:
        return _report_bytes_from_pickle_obj(pickle.loads(data))
    except Exception:
        return _build_docx_bytes(_text_to_paragraphs(data.decode("utf-8", "replace")))


def _is_connected(report_path):
    """None / '' / несуществующий путь / пустая директория → не подключён."""
    if not report_path:
        return False
    try:
        p = Path(report_path)
        if p.is_dir():
            return any(x.is_file() for x in p.iterdir())
        return p.is_file()
    except (TypeError, OSError):
        return False


def main(df: pd.DataFrame = None, report_path: Path = None):
    if df is None:
        df = pd.DataFrame()

    if _is_connected(report_path):
        report_bytes = _load_report_bytes(report_path)          # реальный отчёт
    else:
        report_bytes = _build_docx_bytes(_DEV_REPORT)           # аварийный fallback
        print(f"report_path НЕ подключён — собран fallback .docx ({len(report_bytes)} bytes). "
              f"Подайте реальный отчёт о разработке на report_path.")

    report_dict = {"bin": report_bytes, "ext": ".docx"}
    return {"out_df": df, "report_dict": report_dict}

