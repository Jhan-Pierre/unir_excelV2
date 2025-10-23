# unir_excelV2

Herramienta simple para unir/combinar múltiples archivos Excel en un único archivo de salida. Ideal para consolidar datos de carpetas con hojas similares.

## Requisitos
- Python 3.8+
- pip
- Librerías: pandas, openpyxl (u otras según el formato)

Instalar dependencias:
```
pip install -r requirements.txt
# o
pip install pandas openpyxl
```

## Instalación
1. Clonar el repositorio o copiar los archivos al equipo.
2. Entrar al directorio del proyecto.
3. Instalar dependencias (ver arriba).

## Uso
Ejecutar el script principal (ajustar nombre si es distinto):
```
python main.py --input-dir ./datos --output merged.xlsx
```
Parámetros habituales:
- `--input-dir, -i` : carpeta con archivos `.xlsx`/`.xls`.
- `--output, -o` : nombre/ruta del archivo resultante.
- `--sheet, -s` (opcional): nombre o índice de la hoja a combinar.
- `--ignore-header` (opcional): unir sin duplicar encabezados.

Ejemplos:
```
python main.py -i ./input -o resultado.xlsx
python main.py -i ./input -o resultado.xlsx -s "Hoja1"
```

## Comportamiento esperado
- Lee todos los archivos Excel de la carpeta indicada.
- Concatena filas manteniendo columnas comunes.
- Conserva formato mínimo necesario para exportar a Excel (valores y tipos básicos).
- Sobrescribe el archivo de salida si ya existe (confirmar en la implementación).

## Estructura sugerida del proyecto
- README.md
- requirements.txt
- main.py (o src/)
- tests/ (opcional)
- examples/ (opcional)

## Contribuir
- Abrir un issue para reportar bugs o añadir features.
- Enviar pull requests con cambios pequeños y descriptivos.
- Añadir tests para cambios funcionales importantes.

## Licencia
Licencia MIT (o la que prefiera el autor). Añadir archivo LICENSE si aplica.

## Soporte / Contacto
Abrir un issue en el repositorio con la descripción del problema y pasos para reproducirlo.
