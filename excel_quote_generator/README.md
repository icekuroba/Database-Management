# Excel Quote Generator (VBA)

This macro automates batch quote generation from a source workbook (“cotizador”) and a second workbook containing **quinquennial** values per policy.  
For each policy, it injects the quinquennial value, optionally runs calculation macros, and exports a reduced macro-enabled file with the key sheets.

## Features
- Select source files at runtime (**cotizador** `.xlsm` and **quinquennials** `.xlsx`).
- Optional unprotect of sheets using a list of known passwords.
- Lookup of **quinquennial** by **policy name**.
- Runs dependent macros (`subgrupos`, `Tarifas_enlace`, `Tarifa_Modificaciones`, `resumen`) if present.
- Exports only selected sheets to a timestamped `.xlsm`.
- Creates an output folder automatically under `Documents`.

## Requirements
- Microsoft Excel (with macros enabled).
- Access to:
  - The source quote workbook (`.xlsm`)
  - The quinquennials workbook (`.xlsx`)
- Macros enabled in **Trust Center** settings.

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

