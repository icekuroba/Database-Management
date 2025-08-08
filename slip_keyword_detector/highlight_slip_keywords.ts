function main(workbook: ExcelScript.Workbook) {
  const inicio = new Date().getTime();
  let totalCoincidencias = 0;
  let conteoPalabras: { [clave: string]: number } = {};
  let detallesCoincidencias: { palabra: string; hoja: string; celda: string }[] = [];
  let modoMarcar = false;
  let errorProtegidasMostrado = false;

  // Verifica si se pueden crear hojas
  try {
    const testSheet = workbook.addWorksheet("TEMP_SCRIPT_CHECK");
    testSheet.delete();
  } catch (e) {
    modoMarcar = true;
    console.log("⚠ No se pueden crear hojas, archivo macro o protegido");
  }

  // Lista de productos (anonimizada)
  const productos = [
    { nombre: "Producto A", ramo: "CAT-01", palabras_clave: ["opción", "beneficio", "paquete"] },
    { nombre: "Producto B", ramo: "CAT-02", palabras_clave: ["exceso", "adicional"] },
    { nombre: "Servicio C", ramo: "CAT-03", palabras_clave: ["asistencia", "soporte", "atención"] },
    { nombre: "Cobertura D", ramo: "CAT-04", palabras_clave: ["básico", "ampliado", "premium"] },
    { nombre: "Plan E", ramo: "CAT-05", palabras_clave: ["titular", "dependiente", "familiar"] },
    { nombre: "Complemento F", ramo: "CAT-06", palabras_clave: ["opcional", "extra"] },
    { nombre: "Beneficio G", ramo: "CAT-07", palabras_clave: ["sepelio", "funerario", "asistencia funeraria"] } // ejemplo genérico
  ];

  // Preprocesa claves en minúsculas
  const productosProcesados = productos.map(p => ({
    ...p,
    palabras_clave: p.palabras_clave.map(k => k.toLowerCase())
  }));

  // Crear hojas si es posible
  let resultSheet: ExcelScript.Worksheet | null = null;
  let resumenSheet: ExcelScript.Worksheet | null = null;

  if (!modoMarcar) {
    const existingResult = workbook.getWorksheet("Resultados");
    if (existingResult) existingResult.delete();

    const existingResumen = workbook.getWorksheet("Resumen");
    if (existingResumen) existingResumen.delete();

    try {
      resultSheet = workbook.addWorksheet("Resultados");
      resultSheet.getRange("A1:E1").setValues([["Producto", "Ramo", "Palabra clave", "Hoja", "Celda"]]);
    } catch {}

    try {
      resumenSheet = workbook.addWorksheet("Resumen");
    } catch {}
  }

  const hojas = workbook.getWorksheets();
  let rowIndex = 2;

  for (let sheet of hojas) {
    if (
      sheet.getName() === "Resultados" ||
      sheet.getName() === "Resumen" ||
      sheet.getVisibility() !== ExcelScript.SheetVisibility.visible
    ) continue;

    const range = sheet.getUsedRange();
    if (!range) continue;

    const values = range.getValues();

    const filaOculta: boolean[] = [...Array(range.getRowCount()).keys()].map(i =>
      range.getCell(i, 0).getRow().getHidden()
    );
    const columnaOculta: boolean[] = [...Array(range.getColumnCount()).keys()].map(j =>
      range.getCell(0, j).getColumn().getHidden()
    );

    for (let i = 0; i < values.length; i++) {
      if (filaOculta[i]) continue;

      for (let j = 0; j < values[i].length; j++) {
        if (columnaOculta[j]) continue;

        const rawValue = values[i][j];
        const cellValue = String(rawValue).trim().toLowerCase();
        if (!cellValue) continue;

        const cell = range.getCell(i, j);

        for (let producto of productosProcesados) {
          for (let clave of producto.palabras_clave) {
            if (cellValue.includes(clave)) {
              totalCoincidencias++;
              detallesCoincidencias.push({
                palabra: clave,
                hoja: sheet.getName(),
                celda: cell.getAddress()
              });
              conteoPalabras[clave] = (conteoPalabras[clave] || 0) + 1;

              if (!modoMarcar) {
                try {
                  const textoOriginal = String(cell.getValue());
                  const regex = new RegExp(`\\b(${clave})\\b`, "gi");
                  const textoResaltado = textoOriginal.replace(regex, "🔶$1🔶");
                  cell.setValue(textoResaltado);
                  try {
                    cell.getFormat().getFill().setColor("aqua");
                  } catch {}
                } catch (e) {
                  if (!errorProtegidasMostrado) {
                    errorProtegidasMostrado = true;
                    console.log("⚠ Algunas celdas están protegidas o no se pueden modificar.");
                  }
                }

                if (resultSheet) {
                  const address = cell.getAddressLocal();
                  resultSheet.getCell(rowIndex - 1, 0).setValue(producto.nombre);
                  resultSheet.getCell(rowIndex - 1, 1).setValue(producto.ramo);
                  resultSheet.getCell(rowIndex - 1, 2).setValue(clave);
                  resultSheet.getCell(rowIndex - 1, 3).setValue(sheet.getName());
                  resultSheet.getCell(rowIndex - 1, 4).setValue(address);
                  rowIndex++;
                }
              }
            }
          }
        }
      }
    }
  }

  const fin = new Date().getTime();
  const segundos = Math.round((fin - inicio) / 1000);

  if (!modoMarcar && resumenSheet) {
    resumenSheet.getCell(0, 0).setValue("✅ Script ejecutado correctamente");
    resumenSheet.getCell(1, 0).setValue(`Duración: ${segundos} segundos`);
    resumenSheet.getCell(2, 0).setValue(`Total de palabras encontradas: ${totalCoincidencias}`);
    resumenSheet.getCell(3, 0).setValue("📊 Conteo por palabra clave:");

    let filaResumen = 4;
    for (const clave in conteoPalabras) {
      resumenSheet.getCell(filaResumen, 0).setValue(`${clave}: ${conteoPalabras[clave]}`);
      filaResumen++;
    }
  }

  if (modoMarcar) {
    console.log("Resumen (•ω•)");
    console.log(`Duración: ${segundos} segundos`);
    console.log(`Total: ${totalCoincidencias}`);
    console.log("Conteo por palabra clave:");
    console.log(JSON.stringify(conteoPalabras, null, 2));
    console.log("Detalle de palabras:");
    console.log(JSON.stringify(detallesCoincidencias, null, 2));
  }
}
