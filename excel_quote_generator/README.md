# Excel Quote Generator (VBA)
Este proyecto automatiza la generación masiva de cotizaciones a partir de dos libros de Excel:
un cotizador base y un archivo de quinquenios por póliza.

El proceso que anteriormente se realizaba de forma manual póliza por póliza se transforma en un flujo automático, reduciendo tiempos de ejecución, errores de captura y tareas repetitivas dentro del área administrativa.

## Funcionalidades
- Selección de archivos en tiempo de ejecución (cotizador .xlsm y quinquenios .xlsx).
- Opción de desproteger hojas utilizando una lista de contraseñas conocidas.
- Búsqueda del quinquenio por nombre de póliza.
- Ejecución de macros dependientes (subgrupos, Tarifas_enlace, Tarifa_Modificaciones, resumen) cuando están disponibles.
- Exportación únicamente de las hojas seleccionadas a un archivo .xlsm con marca de tiempo.
- Creación automática de la carpeta de salida dentro de Documentos.

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
- Sistema operativo Windows o IOS
- Microsoft Excel.
- Acceso a:
  - El libro base de cotización (.xlsm)
  - El archivo de quinquenios (.xlsx)
- Macros habilitadas.

## Estructura de los libros
**Libro de cotización (.xlsm)**  
- **POLIZAS**: la columna B contiene los nombres de las pólizas.

**Libro de quinquenios (.xlsx)**  
- Sheet1:  
  - **Columna A **= Nombre de la póliza
  - **Columna B **= Valor del quinquenio (años)

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
- Los nombres de las hojas y las celdas destino pueden modificarse en la sección de constantes del módulo.
- La macro intenta desproteger hojas utilizando un arreglo de contraseñas (vacío por defecto por motivos de seguridad).
-Incluso si ocurre un error durante la ejecución, la configuración de Excel se restaura a su estado original.

## Solución de problemas
- **"No se encontraron las hojas requeridas"** → verificar los nombres o actualizar las constantes.
- **"No se generó salida para una póliza"** → la póliza no existe en la columna B o no se encuentra en el archivo de quinquenios.
- **"No coincide los archivos"**  → Revisar que los archivos seleccionados coincidan

