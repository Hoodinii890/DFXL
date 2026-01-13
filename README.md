# DataFrameXL

## 📦 Instalación / Installation

### Desde GitHub / From GitHub

Puedes instalar la librería directamente desde GitHub usando pip:

```bash
pip install git+https://github.com/Hoodinii890/DFXL.git
```

O desde un branch específico:

```bash
pip install git+https://github.com/Hoodinii890/DFXL.git@main
```

O desde una versión/tag específica:

```bash
pip install git+https://github.com/Hoodinii890/DFXL.git@v0.1.0
```

### Instalación local / Local Installation

Si clonaste el repositorio, puedes instalarlo en modo desarrollo:

```bash
git clone https://github.com/Hoodinii890/DFXL.git
cd DataFrameStyle
pip install -e .
```

### Dependencias / Dependencies

La librería requiere:
- `pandas>=1.3.0`
- `openpyxl>=3.0.0`
- `numpy>=1.20.0`

Estas se instalarán automáticamente al instalar desde GitHub.

---

## 📌 ¿Qué es?
`DataFrameXL` es una extensión de `pandas.DataFrame` que permite trabajar de forma integrada con **Excel (openpyxl)**.
Con este objeto puedes:

- Manipular datos como en cualquier `DataFrame` de pandas.
- Aplicar **estilos de Excel** (fuentes, colores, rellenos, bordes, alineaciones) directamente desde Python.
- Guardar y abrir archivos Excel manteniendo tanto los datos como los estilos.

---

## ⚙️ Cómo funciona
- Al inicializar `DataFrameXL`, se conecta a un archivo Excel (`filename`) y una hoja (`sheet_name`).
- Si el archivo existe, carga los datos y estilos de la hoja en el `DataFrame`.
- Si no existe, crea un nuevo workbook y una hoja vacía.
- Los cambios de datos se sincronizan automáticamente con Excel cuando usas métodos como `setitem`, `loc`, `iloc`, `at`, `iat`.
- Los estilos se almacenan en una estructura interna (`self._styles`) y se aplican al guardar (`save`).
--- 

## 🛠️ Uso de estilos con métodos de pandas
### Asignaciones con setitem (nueva columna o reemplazo)
- Puedes asignar una columna con datos y estilos en un solo paso usando un diccionario con claves data y style.
- Ejemplo conceptual: **`df["A"] = {"data": [1, 2, 3], "style": {...}}`**. Esto registra los datos en pandas y el estilo en `self._styles` para esa columna/fila.

### Asignaciones con at / iat (una sola celda)
- `at` y `iat` aceptan un diccionario con data y style para actualizar una celda y registrar su estilo.
- Ejemplo conceptual: **`df.at[0, "A"] = {"data": 99, "style": {...}}`**. El valor se actualiza en pandas y el estilo se guarda para la celda (fila 0, columna A).

### Asignaciones con loc / iloc (subconjuntos, slices, máscaras)
- `loc` y `iloc` permiten asignar subconjuntos (filas/columnas) con un diccionario **`{"data": ..., "style": ...}`**.
- Para slices o listas de filas, se normalizan los índices afectados y se registra el estilo por cada fila en `self._styles`.
- Con loc también puedes usar condiciones booleanas (p. ej., **`df["A"] > 100`**) para aplicar datos y estilos solo a las filas que cumplan la condición.

## 📖 ¿Qué es `style`?
En `DataFrameXL`, el parámetro `"style": {...}` es un diccionario que describe cómo se debe visualizar una celda, fila, columna o encabezado en Excel.
Este diccionario se guarda en `self._styles` y se aplica a las celdas al momento de guardar (`save()`).

### Cómo se llena el diccionario
El diccionario `style` puede contener las siguientes claves, cada una asociada a un objeto de **`openpyxl.styles`**:

- **`"font"`** → Objeto `openpyxl.styles.Font` Define el tipo de letra, color, tamaño, negrita, cursiva, subrayado, etc.
Ejemplo:
```python
{"font": Font(bold=True, color="FF0000", italic=True)}
```
- **`"fill"`** →  Objeto `openpyxl.styles.PatternFill` Define el color de fondo de la celda.
Ejemplo:
```python
{"fill": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")}
```
- **`"alignment"`** → Objeto `openpyxl.styles.Alignment` Define la alineación horizontal y vertical, ajuste de texto, rotación, etc.
Ejemplo:
```python
{"alignment": Alignment(horizontal="center", vertical="center")}
```
- **`"border"`** → `Objeto openpyxl.styles.Border` Define los bordes de la celda (superior, inferior, izquierdo, derecho).
Ejemplo:
```python
{"border": Border(left=Side(style="thin", color="000000"),
                  right=Side(style="thin", color="000000"),
                  top=Side(style="thin", color="000000"),
                  bottom=Side(style="thin", color="000000"))}
```
Y de esta manera siguiendo la sintaxis y los `openpyxl.styles` se pueden agregar todos los estilos que esta librería disponga.

## Ejemplo completo de `style`
```python
style = {
    "font": Font(bold=True, color="FFFFFF"),
    "fill": PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid"),
    "alignment": Alignment(horizontal="center", vertical="center"),
    "border": Border(left=Side(style="thin", color="FFFFFF"),
                     right=Side(style="thin", color="FFFFFF"))
}
```

## 🛠️ Métodos disponibles
### Métodos de pandas
Todos los métodos de `pandas.DataFrame` están disponibles:

- `loc`, `iloc`, `at`, `iat`
- Operaciones de filtrado, agrupación, merge, etc.
- Cualquier manipulación de datos funciona igual que en pandas.

### Métodos adicionales de estilos
Además, `DataFrameXL` incluye métodos para aplicar estilos sin modificar datos:


- **`set_column_style(col_name, style)`** → Aplica un estilo global a toda la columna.
- **`set_cell_style(row_idx, col_name, style)`** → Aplica un estilo a una celda específica.
- **`set_range_style(row_slice, col_name, style)`** → Aplica un estilo a un rango de filas en una columna.
- **`set_row_style(row_idx, style)`** → Aplica un estilo a toda la fila.
- **`set_header_row_style(style)`** → Aplica un estilo a toda la fila de encabezados.
- **`set_header_cell_style(col_name, style)`** → Aplica un estilo a la celda de encabezado de una columna específica.
- **`set_global_style()`** → Aplica estilos de manera global en todo el documento.
- **`save(filename=None)`** → Aplica los estilos y guarda el archivo Excel. Si no se pasa filename, guarda en el archivo original.
## 📖 Ejemplo de uso
```python
from DFXL import DataFrameXL
from openpyxl.styles import Font, PatternFill

# Crear o abrir un archivo Excel
df = DataFrameXL(filename="reporte.xlsx", sheet_name="hoja1", """df=DataFrame #Si se pasa un DataFrame se contruirá el objeto al rededor de el ignorando la exitencia del archivo filename si existe y cuando se guarde sobrescribirá el archivo.""")

# Modificar datos con pandas
df.loc[0, "A"] = 123

# Aplicar estilos
df.set_cell_style(0, "A", {"font": Font(bold=True, color="FF0000")})
df.set_row_style(1, {"fill": PatternFill("solid", fgColor="FFFF00")})
df.set_header_cell_style("B", {"font": Font(italic=True, color="0000FF")})
```
## 🔄 Reordenamiento con estilos
DataFrameXL ahora soporta los métodos de ordenamiento de pandas con preservación de estilos.
Esto significa que al reordenar filas, los colores, fuentes y formatos aplicados se mueven junto con los datos.

### Métodos disponibles
- `sort_values` → Ordenar por valores de una columna.
- `sort_index` → Ordenar por índice.
- `reindex` → Reordenar filas/columnas según un nuevo índice.
- `reset_index` → Reiniciar el índice manteniendo estilos.
- `sample` → Reordenar filas de manera aleatoria.

### Ejemplo de uso

```python
from DFXL import DataFrameXL

df = DataFrameXL(filename="reporte.xlsx")
df["A"] = {"data": ["Juan", "Ana", "Luis"]}
df["B"] = {"data": [300, 100, 200]}

# Ordenar por columna B y guardar
df_sorted = df.sort_values("B", ignore_index=True)
df_sorted.save("archivo_ordenado.xlsx")
```
👉 Los estilos aplicados a cada fila se mantienen en el archivo `archivo_ordenado.xlsx`.

# Guardar cambios
```python
df.save()
```
## 🎯 Ventajas
- Combina la potencia de pandas con la flexibilidad de openpyxl.
- Permite trabajar con datos y estilos en un solo objeto.
- Facilita la creación de reportes Excel enriquecidos directamente desde Python.
