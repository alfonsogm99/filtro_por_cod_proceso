### README

Previamente, se ha preparado un libro de Excel con las hojas necesarias, tablas dinámicas y datos.


1. Se actualizan los datos en la hoja de origen (bbdd).

**-- `script.py`:**

2. Ejecuta automáticamente 'filtro.py'.

3. Predefine las hojas del Excel.

4. Actualiza las tablas dinámicas.

**-- `filtro.py`:**

5. Filtra los registros según los códigos de operación definidos:

   * Incluyendo únicamente los prefijos válidos.
   * Excluyendo los prefijos excluidos.

6. Separa los datos en dos hojas:

   * BD con los registros válidos.
   * Hoja de depuración con los registros descartados para su revisión (debug).



**Resultado:**

Filtrado de grandes volúmenes de datos (según un código de proceso/ prefijo), generando una base depurada para su revisión. Reduciendo significativamente el tiempo de ejecución y el riesgo de errores manuales.


    -- **Nota:** Por motivos de privacidad y confidencialidad de la información empresarial, el código incluido en este proyecto ha sido adaptado y generalizado. Los nombres de archivos, hojas de cálculo, empresas, columnas y parte de la lógica específica han sido sustituidos por información genérica con el fin de preservar la confidencialidad y evitar la exposición de información interna.
