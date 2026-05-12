import io


def text_from_pdf(file_obj: io.BytesIO) -> str:
    """Extrae texto de un PDF usando pypdf y fallback a pdfplumber."""
    data = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
    parts: list[str] = []

    try:
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(data))
        for page in reader.pages:
            txt = page.extract_text() or ""
            if txt.strip():
                parts.append(txt)
    except Exception:
        parts = []

    if parts:
        return "\n".join(parts)

    try:
        import pdfplumber

        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                if txt.strip():
                    parts.append(txt)
    except Exception:
        parts = []

    return "\n".join(parts)
