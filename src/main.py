import argparse
from pathlib import Path

import helpers.RemoveMetadata as md
from helpers.Logger import mylogger


def get_file_list(folder: Path, extensions: [str]) -> [Path]:
    folder = Path(folder)
    extensions = [ext if ext.startswith('.') else f'.{ext}' for ext in extensions]
    return sorted([f for f in folder.iterdir()
                   if f.suffix.lower() in extensions and not f.name.startswith(('~', '.'))]
                  )

def print_properties(props: dict) -> None:
    for key, value in props.items():
        if value:
            print(f'\t{key:<15}: {value}')


def remove_file_metadata(filename: Path) -> None:
    extension = filename.suffix.lower()

    try:
        mylogger.log.info(f'\n{filename.name.upper()}')
        if extension == '.pdf':
            md.remove_pdf_metadata(filename)
            mylogger.log.info('---> OK')
        elif extension in md.EXTENSIONS:
            md.remove_msoffice_metadata(filename)
            mylogger.log.info('---> OK')
        else:
            mylogger.log.warning('---> Tipo de fichero no soportado.')

    except Exception as e:
        mylogger.log.error(e)

def remove_metadata(path: Path) -> None:
    # Comprobamos que el argumento tiene sentido
    if path.is_file():
        files = [path]
    elif path.is_dir():
        files = get_file_list(path, md.EXTENSIONS)
    else:
        mylogger.log.warning(f"'{path}' no es un fichero ni una carpeta válida.")
        return

    # Comprobamos que hay ficheros para validar
    if not files:
        if path.name == '':
            path = 'el directorio actual.'
        mylogger.log.warning(f"No hay documentos de los tipos soportados en {path}")
        return

    # Hacemos el trabajo
    for filename in files:
        remove_file_metadata(filename)


def display_file_metadata(filename: Path) -> None:
    extension = filename.suffix.lower()

    if extension == '.pdf':
        mylogger.log.info(f'\n{filename.name.upper()}')
        pass
    elif extension in md.EXTENSIONS:
        mylogger.log.info(f'\n{filename.name.upper()}')
        metadata = md.get_msoffice_metadata(filename)
        print_properties(metadata)
    else:
        mylogger.log.warning(f"\nTipo de fichero no soportado: {filename}")

def display_metadata(path: Path) -> None:
    # Comprobamos que el argumento tiene sentido
    if path.is_file():
        files = [path]
    elif path.is_dir():
        files = get_file_list(path, md.EXTENSIONS)
    else:
        mylogger.log.warning(f"'{path}' no es un fichero ni una carpeta válida.")
        return

    # Comprobamos que hay ficheros para validar
    if not files:
        if path.name == '':
            path = 'el directorio actual.'
        mylogger.log.warning(f"No hay documentos de los tipos soportados en {path}")
        return

    # Hacemos el trabajo
    for filename in files:
        display_file_metadata(filename)


def check_file_no_metadata(filename: Path) -> None:
    extension = filename.suffix.lower()

    if extension == '.pdf':
        mylogger.log.info(f'\n{filename.name.upper()}')
        pass
    elif extension in md.EXTENSIONS:
        mylogger.log.info(f'\n{filename.name.upper()}')
        if md.check_msoffice_is_clean(filename):
            mylogger.log.info("---> Limpio")
        else:
            mylogger.log.warning("---> Tiene metadatos")
    else:
        mylogger.log.warning("---> Tipo de fichero no soportado.")

def check_no_metadata(path: Path) -> None:
    # Comprobamos que el argumento tiene sentido
    if path.is_file():
        files = [path]
    elif path.is_dir():
        files = get_file_list(path, md.EXTENSIONS)
    else:
        mylogger.log.warning(f"'{path}' no es un fichero ni una carpeta válida.")
        return

    # Comprobamos que hay ficheros para validar
    if not files:
        if path.name == '':
            path = 'el directorio actual.'
        mylogger.log.warning(f"No hay documentos de los tipos soportados en {path}")
        return

    # Hacemos el trabajo
    for filename in files:
        check_file_no_metadata(filename)


def prepare_argparser():
    parser = argparse.ArgumentParser(description="Borrar metadatos de documentos Office y PDF.")
    parser.add_argument("path", nargs='?', default='.', help="Ruta del office_file o carpeta a procesar.")

    group = parser.add_mutually_exclusive_group()
    group.add_argument("-info", action="store_true", help="Muestra los metadatos sin borrarlos.")
    group.add_argument("-check", action="store_true", help="Comprueba que no hay metadatos.")

    return parser.parse_args()

def main():
    mylogger.reset_log_file()

    args = prepare_argparser()
    path = Path(args.path)

    if args.check:
        check_no_metadata(path)
    elif args.info:
        display_metadata(path)
    else:
        remove_metadata(path)


if __name__ == '__main__':
    main()
