import argparse
from pathlib import Path

import helpers.RemoveMetadata as md
from helpers.Logger import mylogger

VERSION=1.0


def _get_file_list(folder: Path, extensions: [str]) -> [Path]:
    folder = Path(folder)
    extensions = [ext if ext.startswith('.') else f'.{ext}' for ext in extensions]
    return sorted([f for f in folder.iterdir()
                   if f.suffix.lower() in extensions and not f.name.startswith(('~', '.'))]
                  )

def _files_in_path(path:Path) -> list[Path]:
    # Comprobamos que el argumento tiene sentido
    if path.is_file():
        files = [path]
    elif path.is_dir():
        files = _get_file_list(path, md.EXTENSIONS)
    else:
        mylogger.log.warning(f"'{path}' no es un fichero ni una carpeta válida.")
        return []

    # Comprobamos que hay ficheros para validar
    if not files:
        mylogger.log.warning(f"No hay documentos de los tipos soportados en {path.name or 'el directorio actual'}")
        return []

    return files

def _print_properties(props: dict) -> None:
    for key, value in props.items():
        if value:
            mylogger.log.info(f'\t{key:<15}: {value}')


def remove_file_metadata(filename: Path) -> None:
    mylogger.log.info(f'\n{filename.name.upper()}')
    extension = filename.suffix.lower()

    if extension not in md.EXTENSIONS:
        mylogger.log.warning("---> Tipo de fichero no soportado.")
        return

    if extension == '.pdf':
        md.remove_pdf_metadata(filename)
    else:
        md.remove_msoffice_metadata(filename)

    mylogger.log.info('---> OK')

def display_file_metadata(filename: Path) -> None:
    mylogger.log.info(f'\n{filename.name.upper()}')
    extension = filename.suffix.lower()

    if extension not in md.EXTENSIONS:
        mylogger.log.warning("---> Tipo de fichero no soportado.")
        return

    if extension == '.pdf':
        metadata = md.get_pdf_metadata(filename)
    else:
        metadata = md.get_msoffice_metadata(filename)

    _print_properties(metadata)

def check_file_no_metadata(filename: Path) -> None:
    mylogger.log.info(f'\n{filename.name.upper()}')
    extension = filename.suffix.lower()

    if extension not in md.EXTENSIONS:
        mylogger.log.warning("---> Tipo de fichero no soportado.")
        return

    if extension == '.pdf':
        is_clean = md.check_pdf_is_clean(filename)
    else:
        is_clean = md.check_msoffice_is_clean(filename)

    mylogger.log.info("---> Limpio") if is_clean else mylogger.log.warning("---> Tiene metadatos")


def prepare_argparser() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Borrar metadatos de documentos Office y PDF.",
                                     usage="%(prog)s [-h] [-i | --info | -c | --check] [path]")
    parser.add_argument("path", nargs='?', default='.', help="Ruta del office_file o carpeta a procesar.")
    parser.add_argument("-v", "--version", action="version", version=f"%(prog)s v. {VERSION}", help="Muestra la versión del programa.")

    group = parser.add_mutually_exclusive_group()
    group.add_argument("-i", "--info", action="store_true", help="Muestra los metadatos sin borrarlos.")
    group.add_argument("-c", "--check", action="store_true", help="Comprueba que no hay metadatos.")

    return parser.parse_args()


def main():
    mylogger.reset_log_file()
    args = prepare_argparser()

    if args.check:
        operation = check_file_no_metadata
    elif args.info:
        operation = display_file_metadata
    else:
        operation = remove_file_metadata

    path = Path(args.path)
    for file in _files_in_path(path):
        operation(file)


if __name__ == '__main__':
    main()
