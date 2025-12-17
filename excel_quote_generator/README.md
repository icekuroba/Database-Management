# Excel Quote Generator (VBA)
Este proyecto automatiza la generación masiva de cotizaciones a partir de dos libros de Excel:
un cotizador base y un archivo de quinquenios por póliza.

El proceso que anteriormente se realizaba de forma manual póliza por póliza se transforma en un flujo automático, reduciendo tiempos de ejecución, errores de captura y tareas repetitivas dentro del área administrativa.

## Funcionalidades
- Selección de archivos en tiempo de ejecución (cotizador `.xlsm` y quinquenios `.xlsx`).
- Validación de correspondencia entre el archivo de cotizador y el archivo de quinquenios.
- Búsqueda del quinquenio por nombre de póliza.
- Ejecución de macros dependientes (`subgrupo`, `Tarifa`, entre otras) cuando están disponibles.
- Generación de un archivo individual por póliza.
- Creación automática de la carpeta de salida dentro de Documentos.
- Restauración automática del entorno de Excel al finalizar el proceso.

## Estructura del sistema
- Formularios
  - frmCotizador (Interfaz principal)
  - frmSeleccionarArchivos (Seleccionar archivos)
- Módulos
  -  App (Control de flujo)
  - Configuraciones (Validaciones)
  - Quinquenios (Cálculo de censos)
  - Correo (Envio de propuestas)
- Hoja de excel
  - TablaCorreos (Repositorio de datos)

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
1. Abrir una instancia en blanco de Excel.
2. Presionar ALT+F11, ir a Insertar > Módulo y pegar el código desde Module_Cotizador.bas
  (nombre del procedimiento: ProcessPoliciesWithQuinquennials).
3. Guardar el libro de cotización como .xlsm.
4. Cuando el sistema lo solicite, seleccionar:
  - El archivo cotizador (.xlsm)
  - El archivo de quinquenios (.xlsx)
5.La macro creará una carpeta de salida dentro de Documentos y exportará un archivo por cada póliza.

## Notas
- Los nombres de las hojas y parámetros pueden modificarse desde las constantes del código.
- En caso de error, el sistema restaura automáticamente la configuración original de Excel.

## Solución de problemas
- **No se encontraron las hojas requeridas** → verificar los nombres de las hojas.
- **No se generó salida para una póliza** → la póliza no existe en `POLIZARIO` o no se encuentra en el archivo de quinquenios.
- **Los archivos no coinciden** → revisar que el cotizador y el archivo de quinquenios correspondan al mismo mes y mismo contratante.

