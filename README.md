# Eliminador de metadatos en documentos Office y PDF
Este proyecto muestra, chequea y elimina los metadatos en documentos con extensión `.docx`, `.xlsx`, `.pptx` y `.pdf` 
como medida de seguridad antes de publicarlos en Internet o de distribuirlos por cualquier otro medio.

## Características
- **Funciona como programa independiente** (main.py) o como ejecutable Windows (Limpiar.exe).
- **Funciona como módulo importable** (RemoveMetadata) con funciones específicas para chequear y borrar metadatos.
- **Independiente de librerías** para la limpieza de ficheros Word, Excel y Power Point.
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
- `python main.py <fichero_o_carpeta>`: limpia los metadatos de un fichero o de todos los de una carpeta.
- `python main.py <fichero_o_carpeta> -info`: muestra los metadatos sin elminarlos.
- `python main.py <fichero_o_carpeta> -check`: comprueba si los ficheros tienen metadatos.
- Si no se especifica fichero o carpeta, entonces opera sobre el mismo directorio donde está el programa.

## Estructura del proyecto
- `main.py`: Programa principal.
- `RemoveMetadata`: Módulo para manejar metadatos
  - `constants.py`: Definición del tipo de documentos admitidos.
  - `pdf_cleaner.py`: Manejo de metadatos en ficheros PDF
  - `msoffice_cleaner.py`: Manejo de metadatos en ficheros Ms-Office
- `Logger`: Módulo para la gestión de logs por terminal y en fichero

## Contribuciones
Se agradecen las contribuciones mediante fork del repositorio y solicitudes de pull request.

## Licencia
Este proyecto utiliza la Licencia MIT. Consulte el archivo [LICENSE](LICENSE.txt) para más información.

