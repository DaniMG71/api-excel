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
const sheetName = 'KPI_IMs_prueba'; 

// Función para leer Excel
function readExcel() {
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
    const data = readExcel();
    res.json(data);
});

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

    const result = data.find(row => String(row['NUMERO DE TICKET']) === ticketId);

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
