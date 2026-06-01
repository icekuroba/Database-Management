# Excel Quote Generator (VBA)
Este proyecto automatiza la generación masiva de cotizaciones a partir de dos libros de Excel:
un cotizador base y un archivo de quinquenios por póliza.

El proceso que anteriormente se realizaba de forma manual, póliza por póliza, se transforma en un flujo automático, reduciendo tiempos de ejecución, errores de captura y tareas repetitivas dentro del área administrativa.

## Funcionalidades
- Selección de archivos en tiempo de ejecución (cotizador `.xlsm` y quinquenios `.xlsx`).
- Validación de correspondencia entre el archivo de cotizador y el archivo de quinquenios.
- Búsqueda del quinquenio por nombre de póliza.
- Ejecución del flujo de cálculo y llenado de censos.
- Generación de un archivo individual por póliza.
- Creación automática de la carpeta de salida dentro de Documentos.
- Preparación de correos electrónicos con archivos adjuntos.
- Restauración automática del entorno de Excel al finalizar el proceso.

## Estructura del sistema
- Formularios
  - `frmCotizad` (interfaz principal)
  - `frmSeleccionArchivosG` (selección de archivos)
- Módulos
  - `App` (control de flujo)
  - `Configuraciones` (validaciones)
  - `Quinquenio` (cálculo de censos)
  - `Correo` (envío de propuestas)
- Hoja de Excel
  - `TablaCorreos` (repositorio de datos)

## Requisitos
- Sistema operativo **Windows**.
- Microsoft Excel para Windows con soporte para VBA.
- Acceso a:
  - Libro base de cotización (`.xlsm`)
  - Archivo de quinquenios (`.xlsx`)
- Macros habilitadas en el Centro de confianza.

## Estructura de los libros
### Libro de cotización (`.xlsm`)
- `POLIZARIO`: la columna **B** contiene los nombres de las pólizas (a partir de la fila 9).
- `PROPUESTA`: hoja donde se reflejan los resultados del cálculo.

### Libro de quinquenios (`.xlsx`)
- Hoja 1:
  - Columna **A**: nombre de la póliza.
  - Columna **B**: valor del quinquenio (años).

## Uso
1. Descargar el archivo `Select the quote.xlsm`.
2. Guardar el archivo, de preferencia en el Escritorio, junto a la carpeta `image`.
3. Abrir el archivo `.xlsm` en Microsoft Excel.
4. Habilitar macros si Excel lo solicita.
5. Cuando el sistema lo solicite, seleccionar:
   - El archivo cotizador (`.xlsm`).
   - El archivo de quinquenios (`.xlsx`).
6. El programa creará una carpeta de salida dentro de Documentos y exportará un archivo por cada póliza.

## Notas
- Los nombres de las hojas y parámetros pueden modificarse desde las constantes del código.
- Este proyecto funciona como herramienta de apoyo operativo y no sustituye un sistema integral de administración de pólizas.
- En caso de error, el sistema restaura automáticamente la configuración original de Excel.

## Solución de problemas
- **No se encontraron las hojas requeridas** → verificar los nombres de las hojas.
- **No se generó salida para una póliza** → la póliza no existe en `POLIZARIO` o no se encuentra en el archivo de quinquenios.
- **Los archivos no coinciden** → revisar que el cotizador y el archivo de quinquenios correspondan al mismo mes y mismo contratante.
