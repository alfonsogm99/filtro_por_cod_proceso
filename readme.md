### README

Previamente, se ha preparado un libro de Excel con las hojas necesarias, tablas dinámicas y datos.


1. Se actualizan los datos en la hoja de origen del excel '.xlsx'.

**-- `script.py`:**

2. Ejecuta automáticamente 'filtro.py'.

3. Predefine las hojas del libro.

4. Actualiza las tablas dinámicas del informe con la nueva información.

**-- `filtro.py`:**

5. Filtra los registros según los códigos de operación definidos:

   * Incluyendo únicamente los códigos válidos.
   * Excluyendo los códigos específicos configurados.

6. Separa los datos en dos hojas:

   * Base de datos principal con los registros válidos.
   * Hoja de depuración con los registros descartados para su revisión (debug).



**Resultado:**

El proceso automatiza el filtrado de grandes volúmenes de datos, genera una base depurada para el análisis y actualiza automáticamente las tablas dinámicas del informe, reduciendo significativamente el tiempo de ejecución y el riesgo de errores manuales.


-- **Nota:** Por motivos de privacidad y confidencialidad de la información empresarial, el código incluido en este proyecto ha sido adaptado y generalizado. Los nombres de archivos, hojas de cálculo, empresas, columnas y parte de la lógica específica han sido sustituidos por información genérica con el fin de preservar la confidencialidad y evitar la exposición de información interna.