import os
import shutil
import time
import xml.etree.ElementTree as ET
from pathlib import Path

from ..Logger import mylogger


def cambiar_fecha_modificacion(archivo: Path, nueva_fecha: str):
    """ Cambia la fecha de modificación de un archivo. """
    try:
        timestamp = time.mktime(time.strptime(nueva_fecha, "%Y-%m-%d %H:%M:%S"))
        os.utime(archivo, (timestamp, timestamp))
    except Exception as e:
        mylogger.log.error(f"Error cambiando la fecha de modificación: {e}")


def convertir_a_zip(archivo: Path) -> Path|None:
    """ Renombra un archivo de Office a .zip y lo descomprime en un directorio temporal. """
    zip_path = archivo.with_suffix(".zip")
    temp_dir = Path("temp_office")

    try:
        shutil.copy(archivo, zip_path)
        shutil.unpack_archive(zip_path, temp_dir)
    except Exception as e:
        mylogger.log.error(f"Error al convertir el archivo a ZIP: {e}")
        return None
    finally:
        zip_path.unlink()

    return temp_dir


def limpiar_app_properties(app_xml_path: Path):
    """ Elimina las propiedades 'Company' y 'Manager' del archivo app.xml. """
    namespaces = {'ns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}

    try:
        tree = ET.parse(app_xml_path)
        root = tree.getroot()

        for tag in ['Company', 'Manager']:
            element = root.find(f".//ns:{tag}", namespaces)
            if element is not None:
                root.remove(element)

        tree.write(app_xml_path, encoding="utf-8", xml_declaration=True)
    except Exception as e:
        mylogger.log.error(f"Error al limpiar app.xml: {e}")
        raise


def reemplazar_core_properties(core_xml_path: Path):
    """ Reemplaza el archivo core.xml con una versión limpia desde 'helpers/RemoveMetadata/core.xml'. """
    try:
        shutil.copy("helpers/RemoveMetadata/core.xml", core_xml_path)
    except FileNotFoundError:
        mylogger.log.error("Error: No se encontró el archivo helpers/RemoveMetadata/core.xml")
        raise
    except Exception as e:
        mylogger.log.error(f"Error al reemplazar core.xml: {e}")
        raise


def reconstruir_office_file(archivo: Path, temp_dir: Path):
    """ Crea un nuevo archivo de Office con las propiedades eliminadas. """
    try:
        shutil.make_archive(archivo.stem, 'zip', temp_dir)
        shutil.move(f"{archivo.stem}.zip", archivo)
    except Exception as e:
        mylogger.log.error(f"Error al reconstruir el archivo: {e}")
        raise


def remove_properties(archivo: Path) -> bool:
    """ Elimina las propiedades de un archivo de Office sin modificar su contenido principal. """
    temp_dir = convertir_a_zip(archivo)
    if not temp_dir:
        return False

    try:
        # Limpiar app.xml si existe
        app_xml_path = temp_dir / "docProps" / "app.xml"
        if app_xml_path.exists():
            limpiar_app_properties(app_xml_path)

        # Reemplazar core.xml si existe
        core_xml_path = temp_dir / "docProps" / "core.xml"
        if core_xml_path.exists():
            reemplazar_core_properties(core_xml_path)

        # Reensamblar el archivo Office
        reconstruir_office_file(archivo, temp_dir)

        # Restaurar fecha de modificación
        cambiar_fecha_modificacion(archivo, "2000-01-01 01:00:00")
        return True

    except Exception:
        return False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def remove_msoffice_metadata(filename: Path) -> None:
    if remove_properties(filename):
        mylogger.log.info(f"{filename.name} --------> OK")
