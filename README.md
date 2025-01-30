# Eliminador de metadatos en documentos Office y PDF

Este proyecto elimina los metadatos en documentos con extensión `.docx`, `.xlsx`, `.pptx` y `.pdf` como medida de
seguridad antes de publicarlos en Internet o de distribuirlos por cualquier otro medio.

## Características

- **Funciona como programa independiente** en Python (main.py) o como ejecutable Windows (limpia_metadatos.exe).
- **Funciona como módulo importable** (RemoveMetadata) con funciones específicas para mostrar metadatos o borrarlos.
- **Multiplataforma:** Probado en MacOS y en Windows.

## Requisitos
- Python 3.10 o superior
- Librerías necesarias:
  - `pypdf`

Se instalan ejecutando:
```bash
pip install -r requirements.txt
```

# Uso
El programa admite los siguientes argumentos según la operación que se desee:
- `python3 main.py <nombre_de_fichero>`: limpia los metadatos de un fichero.
- `python3 main.py <nombre_de_carpeta>`: limpia los metadatos de los ficheros de una carpeta.

## Estructura del proyecto
- `main.py`: Programa principal.
- `RemoveMetadata`: Módulo para manejar metadatos
  - `constants.py`: Definición de los campos de metadatos y las extensiones de los documentos.
  - `pdf_cleaner.py`: Manejo de metadatos en ficheros PDF
  - `msoffice_cleaner.py`: Manejo de metadatos en ficheros Ms-Office
- `Logger`: Módulo para la gestión de logs por terminal y en fichero

## Contribuciones
Se agradecen las contribuciones mediante fork del repositorio y solicitudes de pull request.

## Licencia
Este proyecto utiliza la Licencia MIT. Consulte el archivo [LICENSE](LICENSE.txt) para más información.

