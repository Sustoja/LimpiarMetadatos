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


def handle_pdf_metadata(filename: Path, delete_metadata: bool) -> None:
    if delete_metadata:
        md.remove_pdf_metadata(filename)
    md.print_pdf_metadata(filename)


def handle_office_metadata(filename: Path, delete_metadata: bool) -> None:
    if delete_metadata:
        md.remove_msoffice_metadata(filename)
    md.print_msoffice_metadata(filename)


def process_file(filename: Path, delete_metadata: bool) -> None:
    extension = filename.suffix.lower()
    if extension == '.pdf':
        handle_pdf_metadata(filename, delete_metadata)
    elif extension in md.EXTENSIONS:
        handle_office_metadata(filename, delete_metadata)
    else:
        mylogger.log.warning(f"Tipo de archivo no soportado: {filename}")


def process_folder(folder: Path, delete_metadata: bool) -> None:
    files = _get_file_list(folder, md.EXTENSIONS)
    if not files:
        mylogger.log.warning(f"No hay documentos de los tipos soportados en: {folder}")
        return

    for filename in files:
        mylogger.log.info(f'\n{filename.name.upper()}')
        process_file(filename, delete_metadata)


def main():
    mylogger.reset_log_file()

    parser = argparse.ArgumentParser(description="Borrar metadatos de documentos Office y PDF.")
    parser.add_argument("path", help="Ruta del archivo o carpeta a procesar.")
    parser.add_argument("--info", action="store_true", help="Muestra los metadatos sin borrarlos.")
    args = parser.parse_args()

    path = Path(args.path)
    delete_metadata = not args.info

    if path.is_file():
        process_file(path, delete_metadata)
    elif path.is_dir():
        process_folder(path, delete_metadata)
    else:
        mylogger.log.warning(f"'{args.path}' no es un archivo ni una carpeta válida.")


if __name__ == '__main__':
    main()
