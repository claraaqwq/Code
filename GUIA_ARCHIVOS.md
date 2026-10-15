# Guía breve de Tracing Validator

## Archivos Python

| Archivo | Para qué sirve | Cómo funciona |
| --- | --- | --- |
| `tracing_validator_FINAL.py` | Es el programa principal y la interfaz que utilizas. | Permite seleccionar archivos, hojas, encabezados y columnas; consultar una cuenta; comparar JE, TB Beginning, TB Ending y Rollforward; y generar el papel de trabajo en Excel. También gestiona configuración, progreso, cancelación y popups. |
| `tracing_sqlite_cache.py` | Guarda los datos localmente para acelerar consultas posteriores. | Importa los Excel por bloques a bases SQLite en la carpeta que elijas. Crea un índice por GLA y recupera los movimientos de cada cuenta sin volver a recorrer el Excel. Si cambia un archivo, crea una nueva importación para ese archivo. |
| `test_tracing_sqlite_cache.py` | Comprueba el funcionamiento de la caché. | Usa archivos temporales para verificar consultas, conservación de valores, reutilización sin reabrir Excel, cambios del archivo original, selección de carpeta y cancelación. |
| `test_tracing_workflow.py` | Comprueba partes del flujo de trabajo. | Verifica que guardar y cargar configuración restaure las selecciones, que una exportación completada guarde el resultado y que cancelar no reemplace un archivo de salida anterior. No prueba visualmente la interfaz. |

## Qué ejecutar

Para usar la aplicación, ejecuta:

```powershell
python tracing_validator_FINAL.py
```

El programa principal importa `tracing_sqlite_cache.py` automáticamente: ambos deben estar en la misma carpeta al ejecutar el código fuente. Los archivos que empiezan con `test_` son para verificar el programa; no necesitas ejecutarlos para trabajar ni incluirlos en el ejecutable distribuido.

Para ejecutar las pruebas:

```powershell
python -m unittest -v test_tracing_sqlite_cache test_tracing_workflow
```

## Flujo de funcionamiento

1. Seleccionas los Excel, sus filas de encabezado y sus hojas.
2. Asignas las columnas de GLA, Business Unit, montos y demás campos necesarios.
3. Eliges una carpeta para la caché local.
4. Preparas la caché con **BUILD / REFRESH CACHE**, o dejas que se cree al generar la primera cuenta.
5. El programa consulta los movimientos de la cuenta, calcula diferencias y genera el Excel de salida.
6. Para otras cuentas, reutiliza las bases SQLite mientras los archivos originales no cambien.

La primera importación puede demorar. Las consultas posteriores evitan recorrer nuevamente los Excel. Los registros de la cuenta seleccionada y el Excel de salida todavía se preparan en memoria.

## Otros archivos y datos guardados

- **`requirements_tracing_validator.txt`:** lista las dependencias externas. Se instalan con `pip install -r requirements_tracing_validator.txt`. SQLite forma parte de Python.
- **Configuración `.json`:** se guarda donde elijas con **SAVE CONFIGURATION** y se recupera con **LOAD CONFIGURATION**. Contiene rutas, hojas, encabezados, mapeos, carpeta de caché, cuenta y tolerancia; no contiene los movimientos del Excel. Los archivos originales deben seguir disponibles en las rutas guardadas.
- **Caché `.sqlite3`:** contiene una copia local de los datos importados y permanece al cerrar la aplicación. No está cifrada por el programa. Puedes eliminarla cuando ya no la necesites; se reconstruirá en una próxima importación.
- **Resultado `.xlsx`:** es el papel de trabajo generado para la cuenta seleccionada.

El código de la aplicación no envía estos datos a servicios externos. La caché se escribe en la ubicación elegida; si seleccionas una carpeta de red o sincronizada, su almacenamiento y sincronización dependen de esa ubicación.
