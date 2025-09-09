// server.js
const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs');
const cors = require('cors');
const path = require('path');
 
const app = express();
app.use(express.json());
app.use(cors());
 
const excelFilePath = path.join(__dirname, 'data.xlsx');
 
 
// Función para leer Excel
function readExcel(sheetName) {
    if (!fs.existsSync(excelFilePath)) {
        return [];
    }
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(worksheet);
}
 
// Función para escribir Excel
function writeExcel(data) {
    const workbook = fs.existsSync(excelFilePath)
        ? xlsx.readFile(excelFilePath)
        : xlsx.utils.book_new();
 
    const worksheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
    xlsx.writeFile(workbook, excelFilePath);
}
// Endpoint para obtener registros
app.get('/tickets', (req, res) => {
    const data = readExcel('KPI_IMs_prueba');
    res.json(data);
});
 
 
// Endpoint para obtener columnas de una hoja específica
app.get('/columns/:sheetName', (req, res) => {
  try {
    const sheetName = req.params.sheetName;
 
    if (!fs.existsSync(excelFilePath)) {
      return res.status(404).json({ error: "El archivo Excel no existe" });
    }
 
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
 
    if (!worksheet) {
      return res.status(404).json({ error: "La hoja no existe en el Excel" });
    }
 
    // Obtener encabezados de la primera fila
    const headers = [];
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = { c: C, r: 0 }; // primera fila
      const cellRef = xlsx.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      headers.push(cell ? cell.v : undefined);
    }
 
    res.json({ columns: headers });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error leyendo columnas del Excel" });
  }
});
 
// Endpoint para agregar un nuevo ticket
app.post('/add-row/tickets', (req, res) => {
  try {
    const newData = req.body;
    const sheetName = 'KPI_IMs_prueba';
 
    // Abrir archivo existente
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
 
    if (!worksheet) {
      return res.status(404).json({ error: "La hoja KPI_IMs_prueba no existe en el Excel" });
    }
 
    // Obtener encabezados de la primera fila
    const headers = [];
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = { c: C, r: 0 };
      const cellRef = xlsx.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      headers.push(cell ? cell.v : undefined);
    }
 
    // Calcular nueva fila
    const newRow = range.e.r + 1; // siguiente fila vacía
    const cellAddress = { c: colIndex, r: newRow }; // fila correcta
 
    // Escribir valores según encabezados
    headers.forEach((header, colIndex) => {
      if (header && newData.hasOwnProperty(header)) {
        const cellAddress = { c: colIndex, r: newRow - 1 };
        const cellRef = xlsx.utils.encode_cell(cellAddress);
        const value = newData[header];
 
        worksheet[cellRef] = {
          t: typeof value === 'number' ? 'n' : 's',
          v: value
        };
      }
    });
 
    // Actualizar rango
    range.e.r = newRow - 1;
    worksheet['!ref'] = xlsx.utils.encode_range(range);
 
    // Guardar archivo
    xlsx.writeFile(workbook, excelFilePath);
 
    res.json({ message: 'Ticket agregado correctamente' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Error agregando ticket' });
  }
});
 
// server.js
 
// Endpoint para obtener todos los tickets con un mismo número
app.get('/tickets/:ticketId', (req, res) => {
  try {
    const ticketId = req.params.ticketId;
    const sheetName = 'KPI_IMs_prueba';
 
    if (!fs.existsSync(excelFilePath)) {
      return res.status(404).json({ error: "Archivo Excel no encontrado" });
    }
 
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      return res.status(404).json({ error: "La hoja KPI_IMs_prueba no existe" });
    }
 
    // Convertimos la hoja a JSON
    const data = xlsx.utils.sheet_to_json(worksheet, { defval: "" }) || [];
 
    // Filtrar por número de ticket
    const tickets = data.filter(row => String(row['NUMERO DE TICKET']) === ticketId);
 
    // Si no encuentra, devolvemos array vacío
    if (tickets.length === 0) {
      return res.json({ message: "No se encontraron tickets con ese número", tickets: [] });
    }
 
    res.json({ tickets });
 
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error buscando tickets" });
  }
});
 
 
// Endpoint para eliminar un ticket por NUMERO DE TICKET
app.delete('/delete-row/tickets/:ticketId', (req, res) => {
  try {
    const ticketId = req.params.ticketId;
    const sheetName = 'KPI_IMs_prueba';
 
    if (!fs.existsSync(excelFilePath)) {
      return res.status(404).json({ error: "Archivo Excel no encontrado" });
    }
 
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      return res.status(404).json({ error: "La hoja KPI_IMs_prueba no existe" });
    }
 
    const data = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
 
    // Buscar el índice de la fila que coincide con el NUMERO DE TICKET
    const rowIndex = data.findIndex(row => String(row['NUMERO DE TICKET']) === ticketId);
    if (rowIndex === -1) {
      return res.status(404).json({ error: "Ticket no encontrado" });
    }
 
    // Eliminar la fila correspondiente
    data.splice(rowIndex, 1);
 
    // Sobrescribir la hoja con los datos restantes
    const newWorksheet = xlsx.utils.json_to_sheet(data);
    workbook.Sheets[sheetName] = newWorksheet;
    xlsx.writeFile(workbook, excelFilePath);
 
    res.json({ message: "Ticket eliminado correctamente" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error eliminando ticket" });
  }
});
 
// Endpoint para actualizar el campo "¿Aplica para el KPI?" de un ticket
app.put('/update-kpi/:ticketId', (req, res) => {
  try {
    const ticketId = req.params.ticketId;
    const { nuevoValor } = req.body;
    const sheetName = 'KPI_IMs_prueba';
 
    if (!fs.existsSync(excelFilePath)) {
      return res.status(404).json({ error: "Archivo Excel no encontrado" });
    }
 
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      return res.status(404).json({ error: `La hoja ${sheetName} no existe` });
    }
 
    // Convertir hoja a JSON
    const data = xlsx.utils.sheet_to_json(worksheet, { defval: "" }) || [];
 
    // Buscar ticket
    const ticket = data.find(row => String(row['NUMERO DE TICKET']) === ticketId);
    if (!ticket) {
      return res.status(404).json({ error: "Ticket no encontrado" });
    }
 
    // Actualizar campo
    ticket["¿Aplica para el KPI?"] = nuevoValor;
 
    // Reescribir hoja
    const newWorksheet = xlsx.utils.json_to_sheet(data);
    workbook.Sheets[sheetName] = newWorksheet;
    xlsx.writeFile(workbook, excelFilePath);
 
    res.json({ message: "KPI actualizado correctamente" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error actualizando KPI" });
  }
});
 
 
// Endpoint para obtener tiendas
app.get('/tiendas', (req, res) => {
    const data = readExcel('DetalleTiendas');
    res.json(data);
});
// Endpoint para agregar un nuevo registro
app.post('/add-row/tiendas', (req, res) => {
  try {
    const newData = req.body; // objeto con los datos a agregar
    const sheetName = 'DetalleTiendas';
    // Abrir archivo existente
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
 
    // Obtener encabezados desde la primera fila
    const headers = [];
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = { c: C, r: 0 }; // primera fila (índice 0)
      const cellRef = xlsx.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      headers.push(cell ? cell.v : undefined);
    }
 
    // Calcular nueva fila
    const lastRow = range.e.r + 1; // última fila con datos (índice base 0 + 1)
    const newRow = lastRow + 1; // siguiente fila vacía
 
    // Escribir valores dinámicamente según encabezados
    headers.forEach((header, colIndex) => {
      if (header && newData.hasOwnProperty(header)) {
        const cellAddress = { c: colIndex, r: newRow - 1 }; // índice base 0
        const cellRef = xlsx.utils.encode_cell(cellAddress);
        const value = newData[header];
 
        // Definir tipo de celda según valor
        worksheet[cellRef] = {
          t: typeof value === 'number' ? 'n' : 's',
          v: value
        };
      }
    });
 
    // Actualizar rango para incluir la nueva fila
    range.e.r = newRow - 1;
    worksheet['!ref'] = xlsx.utils.encode_range(range);
 
    // Guardar archivo
    xlsx.writeFile(workbook, excelFilePath);
 
    res.json({ message: 'Fila agregada correctamente sin tocar fórmulas' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Error agregando la fila' });
  }
});
 
// Endpoint para buscar tiendas por COD SAP o por parte del nombre
app.get('/tiendas/:query', (req, res) => {
  try {
    const rawQuery = req.params.query || "";
    const query = rawQuery.trim().toLowerCase();
    const sheetName = 'DetalleTiendas';
 
    if (!fs.existsSync(excelFilePath)) {
      return res.status(404).json({ error: "Archivo Excel no encontrado" });
    }
 
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      return res.status(404).json({ error: `La hoja ${sheetName} no existe` });
    }
 
    // Convertimos la hoja a JSON (siempre un array)
    const data = xlsx.utils.sheet_to_json(worksheet, { defval: "" }) || [];
 
    // Filtrar por COD SAP exacto o por nombre que contenga la query (case-insensitive)
    const tiendas = data.filter(row => {
      const codSap = String(row['COD SAP'] || "").trim().toLowerCase();
      const nombre = String(row['NOMBRE PTO OPERACIONAL'] || "").trim().toLowerCase();
      return (codSap && codSap === query) || (nombre && nombre.includes(query));
    });
 
    // Devuelve siempre un array (vacío si no hay coincidencias)
    if (tiendas.length === 0) {
      return res.json({ message: "No se encontraron tiendas", tiendas: [] });
    }
 
    res.json({ tiendas });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error buscando tiendas" });
  }
});
 
// Eliminar tienda por COD SAP
app.delete('/tiendas/:codSap', (req, res) => {
  try {
    const codSap = req.params.codSap
    const sheetName = 'DetalleTiendas' // <-- cambiar Tiendas por DetalleTiendas
 
    if (!fs.existsSync(excelFilePath)) {
      return res.status(404).json({ error: "Archivo Excel no encontrado" })
    }
 
    const workbook = xlsx.readFile(excelFilePath)
    const worksheet = workbook.Sheets[sheetName]
 
    if (!worksheet) {
      return res.status(404).json({ error: `La hoja ${sheetName} no existe` })
    }
 
    let data = xlsx.utils.sheet_to_json(worksheet, { defval: "" }) || []
    const index = data.findIndex(row => String(row['COD SAP']).trim() === codSap.trim())
 
    if (index === -1) {
      return res.status(404).json({ error: "No se encontró la tienda" })
    }
 
    data.splice(index, 1) // eliminar la tienda
    workbook.Sheets[sheetName] = xlsx.utils.json_to_sheet(data)
    xlsx.writeFile(workbook, excelFilePath)
 
    res.json({ message: `Tienda con COD SAP ${codSap} eliminada ✅` })
 
  } catch (error) {
    console.error(error)
    res.status(500).json({ error: "Error eliminando la tienda" })
  }
})
 
 
 
// Endpoint para agregar un nuevo registro
app.post('/add-row', (req, res) => {
  try {
    const newData = req.body; // objeto con los datos a agregar
 
    // Abrir archivo existente
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[sheetName];
 
    // Obtener encabezados desde la primera fila
    const headers = [];
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = { c: C, r: 0 }; // primera fila (índice 0)
      const cellRef = xlsx.utils.encode_cell(cellAddress);
      const cell = worksheet[cellRef];
      headers.push(cell ? cell.v : undefined);
    }
 
    // Calcular nueva fila
    const lastRow = range.e.r + 1; // última fila con datos (índice base 0 + 1)
    const newRow = lastRow + 1; // siguiente fila vacía
 
    // Escribir valores dinámicamente según encabezados
    headers.forEach((header, colIndex) => {
      if (header && newData.hasOwnProperty(header)) {
        const cellAddress = { c: colIndex, r: newRow - 1 }; // índice base 0
        const cellRef = xlsx.utils.encode_cell(cellAddress);
        const value = newData[header];
 
        // Definir tipo de celda según valor
        worksheet[cellRef] = {
          t: typeof value === 'number' ? 'n' : 's',
          v: value
        };
      }
    });
 
    // Actualizar rango para incluir la nueva fila
    range.e.r = newRow - 1;
    worksheet['!ref'] = xlsx.utils.encode_range(range);
 
    // Guardar archivo
    xlsx.writeFile(workbook, excelFilePath);
 
    res.json({ message: 'Fila agregada correctamente sin tocar fórmulas' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Error agregando la fila' });
  }
});
 
    // Endpoint para buscar por ticket (ID)
app.get('/tickets/:id', (req, res) => {
  try {
    const ticketId = req.params.id;
    const data = readExcel();
 
    const result = data.filter(row => String(row['NUMERO DE TICKET']) === ticketId);
 
    if (result) {
      res.json(result);
    } else {
      res.status(404).json({ error: 'Ticket no encontrado' });
    }
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Error buscando el ticket' });
  }
});
// Iniciar servidor
const PORT = 3001;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});