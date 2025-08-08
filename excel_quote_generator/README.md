# Excel Quote Generator (VBA)

This macro automates batch quote generation from a source workbook (“cotizador”) and a second workbook containing quinquennial values per policy.  
For each policy, it injects the quinquennial value, runs calculation macros, and exports a reduced macro-enabled file with the key sheets.

## Features
- Select source files at runtime (cotizador `.xlsm` and quinquennials `.xlsx`).
- Unprotect sheets with a list of known passwords (if required).
- Looks up the quinquennial value by policy name.
- Runs dependent macros (subgrupos, tarifas, resumen).
- Exports only selected sheets to a timestamped `.xlsm`.
- Creates an output folder automatically under `Documents`.

## Requirements
- Microsoft Excel (macro-enabled).
- Access to the source quote workbook (`.xlsm`) and the quinquennials workbook (`.xlsx`).
- Macros enabled (Trust Center settings).

## Usage
1. Open a blank Excel instance.
2. Press `ALT+F11`, insert a new Module, and paste the code from `ProcesarPolizasConQuinquenios.bas`.
3. Run `ProcesarPolizasConQuinquenios`.
4. Select the **cotizador** file when prompted.
5. Select the **quinquennials** file when prompted.
6. The macro will create an output folder and export one file per policy.

## Notes
- Sheet names and target cells are configurable in the constants section of the module.
- The macro attempts multiple passwords for protected sheets. Update the list to match your environment.
- Hidden errors are minimized; the routine restores Excel settings even if an error occurs.
