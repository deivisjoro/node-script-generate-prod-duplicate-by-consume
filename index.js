import { fileURLToPath } from 'url';
import { dirname } from 'path';
import fs from 'fs/promises';
import path from 'path';
import xlsx from 'xlsx';
import Papa from 'papaparse';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const directorio = path.join(__dirname, 'fichas');
const directorioSalida = path.join(__dirname, 'bulk_import');
const archivoBOM = path.join(__dirname, 'data', 'bom.xlsx');

function getCellValue(worksheet, rowIndex, colIndex) {
  const cellAddress = { r: rowIndex, c: colIndex };
  const cell = worksheet[xlsx.utils.encode_cell(cellAddress)];
  return cell ? cell.v : null;
}

// Función auxiliar para obtener el valor como string
function obtenerValorString(valor, defaultValue = '') {
  return String(valor || defaultValue).trim();
}

async function leerDirectorio() {
  try {
    // Verificar si el directorio 'fichas' existe
    try {
      await fs.access(directorio);
    } catch (error) {
      throw new Error('El directorio "fichas" no existe.');
    }

    // Verificar y crear el directorio de salida si no existe
    try {
      await fs.access(directorioSalida);
    } catch (error) {
      await fs.mkdir(directorioSalida);
    }

    // Arreglo para almacenar los encabezados acumulados
    const encabezadosAcumulados = [];

    // Arreglo para almacenar todos los detalles acumulados
    const detallesAcumulados = [];

    const procesos = [];
    const pasos = [];

    const archivos = await fs.readdir(directorio);

    for (const archivo of archivos) {
      const rutaCompleta = path.join(directorio, archivo);

      try {
        const stats = await fs.stat(rutaCompleta);

        if (stats.isFile() && path.extname(archivo) === '.xlsx') {
          console.log('Archivo Excel encontrado:', archivo);

          // Leer el archivo Excel
          const workbook = xlsx.readFile(rutaCompleta);
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          // Obtener el encabezado principal
          const headerPrincipal = [];
          for (let rowIndex = 0; rowIndex <= 3; rowIndex++) {
            const cellA = getCellValue(worksheet, rowIndex, 0);
            const cellB = getCellValue(worksheet, rowIndex, 1);

            // Agregar tanto la etiqueta (columna A) como el valor (columna B) al encabezado principal
            headerPrincipal.push({
              label: cellA,
              value: cellB,
            });
          }

          // Agregar los encabezados al arreglo acumulado
          encabezadosAcumulados.push({
            importId: '',
            name: obtenerValorString(headerPrincipal[0].value),
            code: obtenerValorString(headerPrincipal[1].value),
            workshopStockLocation_name: obtenerValorString(headerPrincipal[3].value),
            company_code: 'BASE',
            statusSelect: '3',
            product_code: obtenerValorString(headerPrincipal[2].value)
          });

          // Obtener los detalles
          const ref = worksheet['!ref'];
          const range = xlsx.utils.decode_range(ref);
          const startRow = 6; // Empezar desde la fila 6 para los detalles

          for (let rowIndex = startRow; rowIndex <= range.e.r; rowIndex++) {
            const importId = '';
            const prodProcess_code = obtenerValorString(headerPrincipal[1].value);
            const name = obtenerValorString(getCellValue(worksheet, rowIndex, 0));
            const priority = obtenerValorString(getCellValue(worksheet, rowIndex, 1));
            const workCenter_code = obtenerValorString(getCellValue(worksheet, rowIndex, 2));
            const description = obtenerValorString(getCellValue(worksheet, rowIndex, 3));
            const minCapacityPerCycle = 0;
            const maxCapacityPerCycle = 0;
            const durationPerCycle_seconds = 0;

            // Verificar si al menos una celda de la fila contiene datos antes de agregarla
            if (name || priority || workCenter_code || description) {
              detallesAcumulados.push({
                importId,
                prodProcess_code,
                name,
                priority,
                workCenter_code,
                description,
                minCapacityPerCycle,
                maxCapacityPerCycle,
                durationPerCycle_seconds,
              });
            }
          }
        }
      } catch (error) {
        console.error('Error al procesar el archivo:', error);
      }
    }

    try{
      const workbookBOM = xlsx.readFile(archivoBOM);
      const firstSheetNameBOM = workbookBOM.SheetNames[0];
      const worksheetBOM = workbookBOM.Sheets[firstSheetNameBOM];

      const productos = [];
      let fin = false;
      let colIndex = 5;
      // Objeto para rastrear los códigos y sus conteos
      const conteosPorCodigo = {};
      while(!fin) { 
        const cellAddressCodigo = { r: 1, c: colIndex };
        const cellCodigo = worksheetBOM[xlsx.utils.encode_cell(cellAddressCodigo)];
        const valorCeldaCodigo = cellCodigo ? cellCodigo.v : null;
        
        if (valorCeldaCodigo) {

          const cellAddressNombre = { r: 2, c: colIndex };
          const cellNombre = worksheetBOM[xlsx.utils.encode_cell(cellAddressNombre)];
          const valorCeldaNombre = cellNombre ? cellNombre.v : null;

          // Obtener el conteo actual para este código o establecerlo en 1 si es la primera vez que lo encontramos
          const conteoActual = conteosPorCodigo[valorCeldaCodigo] || 1;

          // Incrementar el conteo para el siguiente elemento del mismo grupo
          conteosPorCodigo[valorCeldaCodigo] = conteoActual + 1;

          productos.push({
            codigo: valorCeldaCodigo,
            consecutivo: `${valorCeldaCodigo}-${conteoActual}`,
            nombre: valorCeldaNombre,
          });
        } else {
          fin = true;
        }

        colIndex++;
      }      

      if(productos){    
        
        productos.forEach(producto=>{
          const fichaEncabezado = encabezadosAcumulados.find(f=>f.product_code==producto.codigo);
          
          procesos.push({
            importId: '',
            name: `PP ${producto.consecutivo} ${producto.nombre}`,
            code: `PP-${producto.consecutivo}`,
            workshopStockLocation_name: fichaEncabezado.workshopStockLocation_name,
            company_code: fichaEncabezado.company_code,
            statusSelect: fichaEncabezado.statusSelect,
            product_code: producto.consecutivo
          });

          //la ficha tiene un codigo (fichaEncabezado.code) y debo buscar en detalles acumulados (detallesAcumulados) la coincidencia con el campo prodProcess_code de detalles acumulados
          const pasosDetalles = detallesAcumulados.filter(d=>d.prodProcess_code==fichaEncabezado.code);

          pasosDetalles.forEach(d=>{
            pasos.push({
              importId: d.importId,
              prodProcess_code: `PP-${producto.consecutivo}`, 
              name: d.name,
              priority: d.priority,
              workCenter_code: d.workCenter_code,
              description: d.description,
              minCapacityPerCycle: d.minCapacityPerCycle,
              maxCapacityPerCycle: d.maxCapacityPerCycle,
              durationPerCycle_seconds: d.durationPerCycle_seconds,
            })
          })
        })
        
      }
    }
    catch(error){
      console.error('Error al procesar el archivo bom.xlsx:', error)
    }

    // Convertir a CSV y escribir el archivo acumulado de encabezados
    const nombreCSVEncabezadosAcumulados = 'production_prodProcess.csv';
    const rutaCSVEncabezadosAcumulados = path.join(directorioSalida, nombreCSVEncabezadosAcumulados);
    const csvDataEncabezadosAcumulados = Papa.unparse(procesos, { delimiter: ';' });
    await fs.writeFile(rutaCSVEncabezadosAcumulados, csvDataEncabezadosAcumulados, 'utf-8');
    console.log('Archivo CSV de Encabezados acumulados creado:', nombreCSVEncabezadosAcumulados);

    // Convertir a CSV y escribir el archivo acumulado de detalles
    const nombreCSVDetallesAcumulados = 'production_prodProcessLine.csv';
    const rutaCSVDetallesAcumulados = path.join(directorioSalida, nombreCSVDetallesAcumulados);
    const csvDataDetallesAcumulados = Papa.unparse(pasos, { delimiter: ';' });
    await fs.writeFile(rutaCSVDetallesAcumulados, csvDataDetallesAcumulados, 'utf-8');
    console.log('Archivo CSV de Detalles acumulados creado:', nombreCSVDetallesAcumulados);

  } catch (error) {
    console.error('Error al leer el directorio:', error.message);
  }
}

leerDirectorio();
