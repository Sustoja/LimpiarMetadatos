from datetime import datetime
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from .constants import META_FIELDS, DATE_FIELDS
from ..Logger import mylogger


def _convert_pdf_datetime(date_str):
    if not date_str:
        return None

    try:
        date_part = date_str[2:-7]
        return datetime.strptime(date_part, "%Y%m%d%H%M%S")
    except ValueError as e:
        mylogger.log.warning(f"Error al convertir la fecha '{date_str}': {e}")
        return None


def _extract_metadata(reader: PdfReader, extension: str) -> {}:
    metadata = {}
    raw_metadata = reader.metadata

    # Campos generales
    metadata.update({
        field: raw_metadata.get(f'/{field}', None) for field in META_FIELDS.get(extension, [])
    })

    # Campos de fecha
    metadata.update({
        field: _convert_pdf_datetime(raw_metadata.get(f'/{field}', None)) for field in DATE_FIELDS.get(extension, [])
    })

    return metadata


def print_pdf_metadata(filename: Path) -> None:
    try:
        reader = PdfReader(filename)
        extension = filename.suffix.lower()

        metadata = _extract_metadata(reader, extension)
        for field, value in metadata.items():
            mylogger.log.info(f"{field}: {value}")
    except Exception as e:
        mylogger.log.error(f"Error al leer los metadatos de {filename}: {e}")


def remove_pdf_metadata(filename: Path) -> bool:
    try:
        reader = PdfReader(filename)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.add_metadata({
            "/Producer": None  # writer.add_metadata({}) no vacía el campo "Producer"
        })

        temp_filename = filename.with_suffix(".temp")
        with open(temp_filename, "wb") as output_file:
            writer.write(output_file)
        temp_filename.replace(filename)
        return True
    except Exception as e:
        mylogger.log.error(f"Error al borrar los metadatos del archivo {filename}: {e}")
        return False
