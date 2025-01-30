import argparse
from pathlib import Path

import helpers.RemoveMetadata as md
from helpers.Logger import mylogger


def _get_file_list(folder: Path, extensions: [str]) -> [Path]:
    folder = Path(folder)
    extensions = [ext if ext.startswith('.') else f'.{ext}' for ext in extensions]
    return sorted([f for f in folder.iterdir()
                   if f.suffix.lower() in extensions and not f.name.startswith(('~', '.'))]
                  )


def process_file(filename: Path) -> None:
    extension = filename.suffix.lower()
    if extension == '.pdf':
        md.remove_pdf_metadata(filename)
    elif extension in md.EXTENSIONS:
        md.remove_msoffice_metadata(filename)
    else:
        mylogger.log.warning(f"Tipo de archivo no soportado: {filename}")


def process_folder(folder: Path) -> None:
    files = _get_file_list(folder, md.EXTENSIONS)
    if not files:
        mylogger.log.warning(f"No hay documentos de los tipos soportados en: {folder}")
        return

    for filename in files:
        process_file(filename)


def main():
    mylogger.reset_log_file()

    parser = argparse.ArgumentParser(description="Borrar metadatos de documentos Office y PDF.")
    parser.add_argument("path", help="Ruta del archivo o carpeta a procesar.")
    args = parser.parse_args()

    path = Path(args.path)

    if path.is_file():
        process_file(path)
    elif path.is_dir():
        process_folder(path)
    else:
        mylogger.log.warning(f"'{args.path}' no es un archivo ni una carpeta válida.")


if __name__ == '__main__':
    main()
