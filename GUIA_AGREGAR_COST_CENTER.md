# Cómo agregar Cost Center como nivel de Rollforward

Esta guía corresponde a la versión actual de `tracing_validator_FINAL.py`. Describe los cambios; todavía no se han aplicado al programa. Busca los nombres de funciones o fragmentos indicados con Ctrl+F, porque los números de línea cambian al editar.

El ejemplo agrega **GL Account + Business Unit + Cost Center** y conserva los niveles existentes. Cost Center se seleccionará como una **columna del Excel** en las cuatro fuentes. No se incluye entrada de un Cost Center fijo/manual.

## 1. Agregar la opción al desplegable

Archivo: `tracing_validator_FINAL.py`.
Función: `_build_ui()`.

Busca `self.rf_level_combo = ttk.Combobox(`. En su parámetro `values`, agrega la opción:

```python
values=[
    "GL Account",
    "GL Account + Business Unit",
    "GL Account + Business Unit + Source",
    "GL Account + Business Unit + Cost Center",
],
```

## 2. Definir las dimensiones de cada nivel

Reemplaza la función `active_dimensions()` completa por:

```python
def active_dimensions(self):
    levels = {
        "GL Account": [],
        "GL Account + Business Unit": ["Business Unit"],
        "GL Account + Business Unit + Source": ["Business Unit", "Source"],
        "GL Account + Business Unit + Cost Center": ["Business Unit", "Cost Center"],
    }
    return levels[self.rf_level_var.get()]
```

El código anterior tenía un `return` final que trataba cualquier opción adicional como Business Unit + Source. Por eso no basta con añadir el texto al desplegable.

Reemplaza también `active_group_helpers()` por:

```python
def active_group_helpers(self):
    helpers = {
        "Business Unit": "__BU",
        "Source": "__SOURCE",
        "Cost Center": "__COST_CENTER",
    }
    return ["__GLA"] + [helpers[d] for d in self.active_dimensions()]
```

`__COST_CENTER` es el nombre interno de la columna normalizada, no un encabezado que debas agregar al Excel original.

## 3. Agregar el campo a los cuatro mapeos

Función: `_build_ui()`.

Reemplaza las cuatro construcciones de `MappingBox` por:

```python
self.je_map = MappingBox(
    maps, "JOURNAL ENTRIES",
    ["GLA", "Business Unit", "Source", "Amount", "Cost Center"])
self.tb_beg_map = MappingBox(
    maps, "TB BEGINNING",
    ["GLA", "Business Unit", "Source", "Amount", "Cost Center"])
self.tb_end_map = MappingBox(
    maps, "TB ENDING",
    ["GLA", "Business Unit", "Source", "Amount", "Cost Center"])
self.rf_map = MappingBox(
    maps, "ROLLFORWARD",
    ["GLA", "Business Unit", "Source", "Beginning", "Variance", "Ending", "Cost Center"])
```

El nuevo campo se añade al final para conservar el orden de los existentes.

### Ajustar las filas de TB para evitar superposiciones

Dentro de `_build_ui()`, busca:

```python
for role in ["Source", "Amount"]:
```

Reemplázalo por:

```python
for role in ["Source", "Amount", "Cost Center"]:
```

Esto deja espacio para el mensaje que aparece debajo de Business Unit. En ambos TB, Cost Center quedará en la fila 5.

Después, en las **cuatro llamadas `.grid(...)` de las etiquetas y entradas Manual Source**, cambia `row=5` por `row=6`. No cambies todas las filas 5 del archivo: solo estas cuatro llamadas.

También ajusta las dos tuplas de Manual Source en `on_rollforward_level_change()`:

```python
(self.tb_beg_manual_source_label, self.tb_beg_manual_source_entry, use_source, 6),
(self.tb_end_manual_source_label, self.tb_end_manual_source_entry, use_source, 6),
```

## 4. Mostrar u ocultar Cost Center según el nivel

Función: `on_rollforward_level_change()`.

Después de:

```python
use_source = "Source" in dims
```

agrega:

```python
use_cost_center = "Cost Center" in dims
```

Dentro del bucle que recorre los cuatro mapeos, agrega la última línea:

```python
for box in [self.je_map, self.tb_beg_map, self.tb_end_map, self.rf_map]:
    box.show_role("Business Unit", use_bu)
    box.show_role("Source", use_source)
    box.show_role("Cost Center", use_cost_center)
```

Conserva el resto de la función, incluidos los mensajes de Business Unit y Manual Source.

## 5. Validar que los TB tengan la columna

Función: `validate_mapping()`.

JE y Rollforward ya recorren `dims`, así que validarán Cost Center automáticamente. En el bucle de TB Beginning y TB Ending, después de la validación de Source, agrega:

```python
if "Cost Center" in dims:
    self.validate_column(box, "Cost Center", cols, name)
```

Usa la misma indentación que `if "Source" in dims and not manual_source:`. Así se ejecutará para cada TB.

## 6. Enviar el nuevo mapeo al proceso de generación

Función: `generate()`.

Dentro de `self.worker_config`, reemplaza las cuatro entradas de mapeo por:

```python
"je": {r: self.je_map.get(r) for r in ["GLA", "Business Unit", "Source", "Amount", "Cost Center"]},
"beg": {r: self.tb_beg_map.get(r) for r in ["GLA", "Business Unit", "Source", "Amount", "Cost Center"]},
"end": {r: self.tb_end_map.get(r) for r in ["GLA", "Business Unit", "Source", "Amount", "Cost Center"]},
"rf": {r: self.rf_map.get(r) for r in ["GLA", "Business Unit", "Source", "Beginning", "Variance", "Ending", "Cost Center"]},
```

No reemplaces el diccionario completo: conserva nivel, dimensiones, valores manuales y filas de encabezado.

`build_cache()` no necesita este campo en su configuración: la importación SQLite conserva todas las columnas y utiliza GLA como índice.

## 7. Incorporarlo al cálculo de la comparación

Función: `_generate_worker()`.

### 7.1 Preparar JE y ambos TB

En la función interna `prepare_detail()`, antes de asignar `out["__AMOUNT"]`, agrega:

```python
if "Cost Center" in dims:
    out["__COST_CENTER"] = norm_key(out[mapping["Cost Center"]])
```

Esto afecta a JE, TB Beginning y TB Ending porque los tres utilizan esa función interna.

### 7.2 Preparar Rollforward

Después de la asignación de `rf["__SOURCE"]`, agrega al mismo nivel que los otros `if`:

```python
if "Cost Center" in dims:
    rf["__COST_CENTER"] = norm_key(rf[cfg["rf"]["Cost Center"]])
```

### 7.3 Agrupar por Cost Center

Busca el bloque que comienza con `group_cols = ["__GLA"]`. Después de los bloques de Business Unit y Source, agrega:

```python
if "Cost Center" in dims:
    group_cols.append("__COST_CENTER")
```

Los `groupby` y `merge` ya utilizan `group_cols`; no hay que modificar cada uno.

### 7.4 Darle su nombre visible en la validación

En el bloque de `rename_map`, antes de `val.rename(...)`, agrega:

```python
if "Cost Center" in dims:
    rename_map["__COST_CENTER"] = "Cost Center"
```

### 7.5 Retirar las columnas internas de los detalles exportados

Reemplaza `helper_cols` por:

```python
helper_cols = ["__GLA", "__BU", "__SOURCE", "__COST_CENTER", "__AMOUNT"]
```

En `rf_original = rf.drop(...)`, cambia su lista `columns` a:

```python
columns=["__GLA", "__BU", "__SOURCE", "__COST_CENTER", "__BEG", "__VAR", "__END"],
```

Esto elimina únicamente los auxiliares. La columna original de Cost Center sigue formando parte del detalle.

## 8. Incluirlo en el Excel de salida

Función: `_write_working_paper()`.

Después de:

```python
if "Source" in cfg["dimensions"]:
    validation_cols.append("Source")
```

agrega:

```python
if "Cost Center" in cfg["dimensions"]:
    validation_cols.append("Cost Center")
```

También reemplaza el título fijo:

```python
row = section_title(row, "VALIDATION BY GLA + BUSINESS UNIT")
```

por uno que corresponda al nivel elegido:

```python
row = section_title(row, f"VALIDATION BY {cfg['level'].upper()}")
```

No necesitas modificar `write_working_paper()`: esa función gestiona el archivo temporal y el reemplazo seguro del destino, no las columnas.

## 9. Permitir cargar configuraciones antiguas

`_configuration()` ya guarda todos los campos de cada mapeo, así que guardará Cost Center automáticamente.

Pero **sí debes modificar `load_configuration()`** para que acepte JSON antiguos que todavía no tengan ese campo.

En la validación de los mapeos, cambia:

```python
if not isinstance(cfg["mappings"][k][role], str):
```

por:

```python
if not isinstance(cfg["mappings"][k].get(role, ""), str):
```

En la restauración de los campos, cambia:

```python
variable.set(cfg["mappings"][k][role])
```

por:

```python
variable.set(cfg["mappings"][k].get(role, ""))
```

Así las configuraciones anteriores se cargarán con Cost Center vacío. Si después eliges el nuevo nivel, deberás mapear esa columna antes de generar.

## 10. Verificar con un ejemplo pequeño

Usa una sola cuenta, una Business Unit y dos Cost Centers. Los archivos pueden tener otros nombres de encabezado: selecciónalos en los desplegables correspondientes.

| GLA | Business Unit | Cost Center | TB Beginning | JE | TB Ending |
| --- | --- | --- | ---: | ---: | ---: |
| 100 | BU1 | CC01 | 100 | 20 | 120 |
| 100 | BU1 | CC02 | 200 | -10 | 190 |

En Rollforward, agrega dos filas para la cuenta 100 y BU1: CC01 con Beginning 100, Variance 20 y Ending 120; CC02 con Beginning 200, Variance -10 y Ending 190.

Resultado esperado con el nuevo nivel: **dos filas de validación**, una por Cost Center, con diferencias cero y resultados PASS.

Después comprueba que:

1. Al elegir un nivel anterior, Cost Center se oculte y la aplicación siga funcionando.
2. Si falta el mapeo de Cost Center con el nuevo nivel activo, aparezca un error de validación.
3. Guardar y cargar configuración conserve el mapeo de Cost Center.
4. Un JSON antiguo se cargue dejando ese campo vacío.
5. Si aumentas TB Ending de CC01 en 5 y reduces CC02 en 5, ambos fallen por separado al comparar contra el Rollforward sin modificar. Esto verifica que realmente se esté agrupando por Cost Center, aunque el total de la cuenta siga siendo el mismo.

Ejecuta también las pruebas existentes:

```powershell
python -m unittest -v test_tracing_sqlite_cache test_tracing_workflow
```

En `test_tracing_workflow.py`, amplía la lista `roles` del ejemplo de configuración para incluir `"Cost Center"` y comprueba que su mapeo se restaure. Ese ejemplo usa objetos simulados: las pruebas actuales no verifican por sí solas el cálculo del nuevo nivel; agrega una prueba del caso anterior para cubrirlo.

## Qué archivos no necesitan cambios para este ejemplo

- `tracing_sqlite_cache.py`: ya almacena todas las columnas y sigue consultando por GLA. Una caché existente sirve si la columna ya estaba en el Excel al importarlo; si agregas la columna al archivo original, se detectará el cambio y se volverá a importar.
- `requirements_tracing_validator.txt`: no se agregan dependencias.
- `test_tracing_sqlite_cache.py`: el mecanismo de almacenamiento no cambia.
- `load_data_for_generation()`: ya devuelve todas las columnas de la cuenta seleccionada.
- `_prepared()`: no se utiliza en el flujo actual de generación. La preparación que debes modificar está en la función interna `prepare_detail()` de `_generate_worker()`.
