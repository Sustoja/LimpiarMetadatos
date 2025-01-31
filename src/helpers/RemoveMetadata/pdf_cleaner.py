from datetime import datetime
from pathlib import Path
from pypdf import PdfReader, PdfWriter


META_FIELDS = ['Author', 'Creator', 'Keywords', 'Producer', "Subject", "Title"]
DATE_FIELDS = ["CreationDate", "ModDate"]
MANDATORY_KEYS = ['Producer']


def _convert_pdf_datetime(date_str) -> datetime | None:
    if date_str is None:
        return None

    try:
        date_part = date_str[2:-7]
        return datetime.strptime(date_part, "%Y%m%d%H%M%S")
    except Exception:
        raise RuntimeError(f"Error al convertir la fecha '{date_str}'")

def _extract_metadata(reader: PdfReader, extension: str) -> {}:
    metadata = {}
    raw_metadata = reader.metadata

    # Campos generales
    metadata.update({
        field: raw_metadata.get(f'/{field}', None) for field in META_FIELDS
    })

    # Campos de fecha
    metadata.update({
        field: _convert_pdf_datetime(raw_metadata.get(f'/{field}', None)) for field in DATE_FIELDS
    })

    return metadata

def remove_pdf_metadata(filename: Path) -> None:
    temp_filename = filename.with_suffix(".temp")

    try:
        reader = PdfReader(filename)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.add_metadata({
            "/Producer": None  # writer.add_metadata({}) no vacía el campo "Producer"
        })

        with open(temp_filename, "wb") as output_file:
            writer.write(output_file)
        temp_filename.replace(filename)

    except Exception as e:
        raise RuntimeError(f'Error al borrar los metadatos: {e}')

    finally:
        temp_filename.unlink(missing_ok=True)

def get_pdf_metadata(filename: Path) -> dict | None:
    try:
        reader = PdfReader(filename)
        extension = filename.suffix.lower()
        metadata = _extract_metadata(reader, extension)
    except Exception:
        raise RuntimeError("No se ha podido extraer los metadatos.")

    return metadata

def check_pdf_is_clean(filename: Path) -> bool:
    metadata = get_pdf_metadata(filename)
    ok = True
    for key, value in metadata.items():
        if not key in MANDATORY_KEYS and value:
            ok = False
            break

    return ok
