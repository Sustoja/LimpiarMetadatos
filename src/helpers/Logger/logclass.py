import colorama
import logging
import re
from colorama import Fore, Style
from pathlib import Path

colorama.init(autoreset=True)


class TerminalFilter(logging.Filter):
    """
    Filtro para colorear el texto en el terminal según el nivel de criticidad mensajes.
    """
    def filter(self, record):
        color = Fore.RED if record.levelno > logging.INFO else Fore.WHITE
        record.msg = f'{color}{record.msg}'
        return True # Permite que el mensaje continúe al manejador

class FileFilter(logging.Filter):
    """
    Filtro para eliminar los códigos de color  antes de escribirlos en el archivo de registro.
    """
    def filter(self, record):
        color_codes = re.compile(r'\x1B\[[0-9;]*[mK]')
        record.msg = color_codes.sub('', record.msg)
        return True  # Permite que el mensaje continúe al manejador

class MyLogger:
    """
    Logger con soporte para terminal coloreado y escritura en archivo.
    """
    LOG_FILE_NAME = 'eventos.log'

    def __init__(self, log_level = logging.DEBUG):
        self.log = logging.getLogger(__name__)
        self._setup_handlers()
        self.log.setLevel(log_level)

    @property
    def log_path(self) -> Path|None:
        for handler in self.log.handlers:
            if hasattr(handler, "baseFilename"):
                return Path(getattr(handler, 'baseFilename'))
        return None

    @property
    def log_size(self) -> int:
        return self.log_path.stat().st_size if self.log_path and self.log_path.exists() else 0

    def _setup_handlers(self):
        # Manejador para el terminal
        stream_handler = logging.StreamHandler()
        stream_handler.setLevel(logging.DEBUG)
        stream_handler.setFormatter(logging.Formatter('%(message)s'))
        stream_handler.addFilter(TerminalFilter())

        # Manejador para el archivo
        file_handler = logging.FileHandler(self.LOG_FILE_NAME)
        file_handler.setLevel(logging.DEBUG)
        file_handler.addFilter(FileFilter())

        # Agregar manejadores al logger
        self.log.addHandler(stream_handler)
        self.log.addHandler(file_handler)

    def reset_log_file(self, limit=-1) -> None:
        if self.log_size > limit:
            with self.log_path.open('w') as log_file:
                log_file.truncate(0)
