import os
import shutil
import time
import xml.etree.ElementTree as ET
from pathlib import Path

EXTENDED_PROPERTIES = ['Company', 'Manager']
MANDATORY_KEYS = ['revision', "created", "modified"]

def _parse_xml_to_dict(file_path: Path) -> dict:
    tree = ET.parse(file_path)
    root = tree.getroot()

    data = {}

    for elem in root:
        tag = elem.tag.split('}')[-1]
        data[tag] = elem.text if elem.text is not None else ""

    return data


def _cambiar_fecha_modificacion(archivo: Path, nueva_fecha: str) -> None:
    try:
        timestamp = time.mktime(time.strptime(nueva_fecha, "%Y-%m-%d %H:%M:%S"))
        os.utime(archivo, (timestamp, timestamp))
    except Exception:
        raise RuntimeError("Error al cambiar la fecha de modificación.")


def _descomprimir_fichero_office(archivo_office: Path) -> Path | None:
    zip_file = archivo_office.with_suffix(".zip")
    temp_dir = Path("temp_office")

    try:
        shutil.copy(archivo_office, zip_file)
        shutil.unpack_archive(zip_file, temp_dir)
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise RuntimeError("Error al descomprimir el fichero.")
    finally:
        zip_file.unlink()

    return temp_dir


def _borrar_extended_properties(unzipped_dir: Path) -> None:
    """ Elimina las propiedades 'Company' y 'Manager' que están en app.xml. """

    try:
        app_xml_path = unzipped_dir / "docProps" / "app.xml"
        namespaces = {'ns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}

        tree = ET.parse(app_xml_path)
        root = tree.getroot()

        for tag in EXTENDED_PROPERTIES:
            element = root.find(f".//ns:{tag}", namespaces)
            if element is not None:
                root.remove(element)

        tree.write(app_xml_path, encoding="utf-8", xml_declaration=True)

    except Exception:
        raise RuntimeError("Error al borrar las propiedades extendidas.")


def _borrar_core_properties(unzipped_dir: Path) -> None:
    """ Elimina las propiedades principales que se encuentran en core.xml. """

    empty_metadata = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" 
    xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title></dc:title><dc:subject></dc:subject><dc:creator></dc:creator><cp:keywords></cp:keywords>
    <dc:description></dc:description><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>1</cp:revision>
    <dcterms:created xsi:type="dcterms:W3CDTF">2000-01-01T00:00:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2000-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>'''

    try:
        core_xml_path = unzipped_dir / "docProps" / "core.xml"
        with open(core_xml_path, 'w') as fich:
            fich.write(empty_metadata)
    except Exception:
        raise RuntimeError("Error al limpiar las propiedades principales.")


def _reconstruir_fichero_office(archivo_office: Path, temp_dir: Path):
    try:
        shutil.make_archive(archivo_office.stem, 'zip', temp_dir)
        shutil.move(f"{archivo_office.stem}.zip", archivo_office)
    except Exception:
        raise RuntimeError("Error al reconstruir el fichero tras limpiarlo.")


def remove_msoffice_metadata(archivo_office: Path) -> None:
    random_new_date = "2000-01-01 01:00:00"

    unzipped_dir = _descomprimir_fichero_office(archivo_office)

    try:
        _borrar_core_properties(unzipped_dir)
        _borrar_extended_properties(unzipped_dir)
        _reconstruir_fichero_office(archivo_office, unzipped_dir)
        _cambiar_fecha_modificacion(archivo_office, random_new_date)
    except Exception as e:
        raise RuntimeError(e)
    finally:
        shutil.rmtree(unzipped_dir, ignore_errors=True)


def get_msoffice_metadata(archivo_office: Path) -> dict | None:
    unzipped_dir = _descomprimir_fichero_office(archivo_office)

    try:
        core = _parse_xml_to_dict(unzipped_dir / "docProps" / "core.xml")
        app =_parse_xml_to_dict(unzipped_dir / "docProps" / "app.xml")
        core.update({item: app[item] for item in EXTENDED_PROPERTIES if app.get(item)})
    except Exception as e:
        raise RuntimeError(e)
    finally:
        shutil.rmtree(unzipped_dir, ignore_errors=True)

    return core


def check_msoffice_is_clean(archivo_office: Path) -> bool:
    metadata = get_msoffice_metadata(archivo_office)
    ok = True
    for key, value in metadata.items():
        if not key in MANDATORY_KEYS and value:
            ok = False
            break

    return ok