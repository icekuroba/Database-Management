# Excel Quote Generator (VBA)

Este programa automatiza la generación masiva de cotizaciones a partir de dos libros el primero cotizador y un segundo quinquenios por póliza.
Donde el proceso de pólizas solia ser manual y la finalidad de este sistema es que sea automatica.

## Funcionalidades
- Selección de archivos en tiempo de ejecución (cotizador .xlsm y quinquenios .xlsx).
- Opción de desproteger hojas utilizando una lista de contraseñas conocidas.
- Búsqueda del quinquenio por nombre de póliza.
- Ejecución de macros dependientes (subgrupos, Tarifas_enlace, Tarifa_Modificaciones, resumen) cuando están disponibles.
- Exportación únicamente de las hojas seleccionadas a un archivo .xlsm con marca de tiempo.
- Creación automática de la carpeta de salida dentro de Documentos.

## Requisitos
- Microsoft Excel (con macros habilitadas).
- Acceso a:
  - El libro base de cotización (.xlsm)
  - El archivo de quinquenios (.xlsx)
- Macros habilitadas desde la configuración del Centro de confianza.

## Workbook Layout

**Quote workbook (.xlsm)**  
- `POLICIES`: Column **B** contains policy names, starting from **row 9**.  
- `RENEWAL_PROPOSAL`: Target cell for quinquennial = **D15**.  
- Optional: `TEXTS` and `ENDORSEMENTS` (copied if present).

**Quinquennials workbook (.xlsx)**  
- Sheet1:  
  - **Column A** = Policy name  
  - **Column B** = Quinquennial value (years)

## Usage
1. Open a blank Excel instance.
2. Press `ALT+F11`, go to **Insert > Module**, and paste the code from `Module_Cotizador.bas`  
   *(procedure name: `ProcessPoliciesWithQuinquennials`)*.
3. Save the quote workbook as `.xlsm`.
4. Run `ProcessPoliciesWithQuinquennials` via `ALT+F8`.
5. When prompted, select:  
   - The **cotizador** file (`.xlsm`)  
   - The **quinquennials** file (`.xlsx`)
6. The macro will create an output folder under your **Documents** and export one file per policy.

## Notes
- Sheet names and target cells can be changed in the constants section of the module.
- The macro tries multiple passwords for protected sheets (array is empty by default for security).
- Even if an error occurs, Excel settings are restored to their original state.

## Troubleshooting
- **"Required sheets not found"** → check names or update constants.  
- **"No output per policy"** → policy missing in column B or not found in quinquennials file.  
- **Protected sheet could not be unprotected** → add the real password to the `passwords` array.

