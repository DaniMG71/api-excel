// models/DynamicTicketsProblem.js
const { DataTypes } = require("sequelize");
const sequelize = require('../config/database'); // Asegúrate de que tu config/database esté correcto

async function getTicketsProblemModel() {
  // 1️⃣ Obtener columnas desde la tabla
  const [columns] = await sequelize.query(`
    SELECT column_name AS "COLUMN_NAME", data_type AS "DATA_TYPE"
    FROM information_schema.columns
    WHERE LOWER(table_name) = 'ticketsproblem'
      AND table_schema = 'public';
  `);

  if (!columns.length) throw new Error("No se encontraron columnas en TicketsProblem");

  const modelDefinition = {};
  let primaryKey = null;

  const sqlToSequelize = {
    bigint: DataTypes.BIGINT,
    integer: DataTypes.INTEGER,
    "character varying": DataTypes.STRING,
    text: DataTypes.TEXT,
    date: DataTypes.DATE,
    time: DataTypes.TIME,
    "timestamp without time zone": DataTypes.DATE,
    "timestamp with time zone": DataTypes.DATE,
    boolean: DataTypes.BOOLEAN,
    numeric: DataTypes.DECIMAL,
    bit: DataTypes.TEXT,
    decimal: DataTypes.DECIMAL,
    double: DataTypes.DOUBLE,
  };

  for (const col of columns) {
  const name = col.COLUMN_NAME;
  const type =
    sqlToSequelize[col.DATA_TYPE.toLowerCase()] || DataTypes.STRING;

  modelDefinition[name] = { 
    type,
    field: name  // 👈 Usa el nombre EXACTO de PostgreSQL
  };

  if (name.toUpperCase() === "NUMERO_TICKET") {
    modelDefinition[name].primaryKey = true;
    modelDefinition[name].allowNull = false;
  }
}


  return sequelize.define("TicketsProblem", modelDefinition, {
    tableName: "ticketsproblem", 
    timestamps: false,
    createdAt: false,
    updatedAt: false,
  });
}

module.exports = getTicketsProblemModel;