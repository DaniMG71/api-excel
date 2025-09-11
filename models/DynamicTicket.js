// models/DynamicTicket.js
const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

async function getTicketModel() {
  // Traer información de las columnas de la tabla
  const [columns] = await sequelize.query(`
    SELECT COLUMN_NAME, DATA_TYPE
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = 'Tickets'
  `);

  if (!columns.length) throw new Error('No se encontraron columnas en Tickets');

  const modelDefinition = {};
  let primaryKey = null;

  // Mapear tipos de SQL Server a Sequelize
  const sqlToSequelize = {
    bigint: DataTypes.BIGINT,
    int: DataTypes.INTEGER,
    varchar: DataTypes.STRING,
    nvarchar: DataTypes.STRING,
    text: DataTypes.TEXT,
    date: DataTypes.DATE,
    time: DataTypes.TIME,
    datetime: DataTypes.DATE,
    bit: DataTypes.BOOLEAN,
    decimal: DataTypes.DECIMAL,
    float: DataTypes.FLOAT,
  };

  for (const col of columns) {
    const name = col.COLUMN_NAME;
    const type = sqlToSequelize[col.DATA_TYPE.toLowerCase()] || DataTypes.STRING;

    modelDefinition[name] = { type };

    // Detectar primary key (ejemplo: NUMERO_DE_TICKET)
    if (name.toUpperCase().includes('NUMERO') && name.toUpperCase().includes('TICKET')) {
      primaryKey = name;
      modelDefinition[name].primaryKey = true;
    }
  }

  return sequelize.define('Ticket', modelDefinition, {
    tableName: 'Tickets',
    timestamps: false,
    // Sequelize no agregará columna id automáticamente
    createdAt: false,
    updatedAt: false,
  });
}

module.exports = getTicketModel;
