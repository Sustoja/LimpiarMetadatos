import logging
import os
from pathlib import Path


# Colores para el texto por pantalla.
COLORS = {
    'WHITE': '\033[97m',
    'RED': '\033[91m',
    'GREEN': '\033[92m',
}
DEFAULT_COLOR = COLORS['WHITE']

def is_color_supported() -> bool:
    return os.name == 'posix'

class TerminalFilter(logging.Filter):
    """
    Filtro para colorear el texto en el terminal según el nivel de criticidad mensajes.
    Compatible solo con sistemas que soportan colores ANSI (no en Windows)
    """
    def filter(self, record):
        if is_color_supported():
            color = COLORS["RED"] if record.levelno > logging.INFO else DEFAULT_COLOR
            record.msg = f'{color}{record.msg}{DEFAULT_COLOR}'
        return True # Permite que el mensaje continúe al manejador

class FileFilter(logging.Filter):
    """
    Filtro para eliminar los códigos de color  antes de escribirlos en el archivo de registro.
    """
    def filter(self, record):
        for color in COLORS.values():
            record.msg = record.msg.replace(color, "")
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
