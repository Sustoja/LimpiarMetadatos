import logging
from .logclass import MyLogger

# Configuración del logger global que usarán todos los módulos que importen este paquete.

# Constantes para el nivel mínimo de logging y el nombre del fichero de logging
DEFAULT_LOG_LEVEL = logging.DEBUG
LOG_FILE_NAME = 'eventos.log'


try:
    logger = MyLogger(LOG_FILE_NAME, DEFAULT_LOG_LEVEL)
except Exception as e:
    raise RuntimeError(f"Error al inicializar el logger global: {e}")

__all__ = ["logger", "LOG_FILE_NAME"]
