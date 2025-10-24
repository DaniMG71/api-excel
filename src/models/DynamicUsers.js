// models/DynamicUser.js
const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

async function getUserModel() {
  // Traer información de las columnas de la tabla
  const [columns] = await sequelize.query(`
    SELECT column_name AS "COLUMN_NAME", data_type AS "DATA_TYPE"
    FROM information_schema.columns
    WHERE table_name = 'users'
      AND table_schema = 'public'
  `);

  if (!columns.length) throw new Error('No se encontraron columnas en Users');

  const modelDefinition = {};

  // Mapear tipos PostgreSQL a Sequelize
  const sqlToSequelize = {
    bigint: DataTypes.BIGINT,
    integer: DataTypes.INTEGER,
    "character varying": DataTypes.STRING,
    text: DataTypes.TEXT,
    date: DataTypes.DATEONLY, // solo fecha
    time: DataTypes.TIME,
    "timestamp without time zone": DataTypes.DATE,
    "timestamp with time zone": DataTypes.DATE,
    boolean: DataTypes.BOOLEAN,
    bit: DataTypes.BOOLEAN, // mapear bit a boolean
    numeric: DataTypes.DECIMAL,
    "double precision": DataTypes.FLOAT,
  };

  for (const col of columns) {
    const name = col.COLUMN_NAME.toLowerCase(); // usar minúsculas
    const type = sqlToSequelize[col.DATA_TYPE.toLowerCase()] || DataTypes.STRING;

    modelDefinition[name] = { type };

    // Detectar primary key por nombre exacto
    if (name === 'id') {
      modelDefinition[name].primaryKey = true;
      modelDefinition[name].autoIncrement = true;
      modelDefinition[name].allowNull = false;
    }
  }

  return sequelize.define('User', modelDefinition, {
    tableName: 'users',
    timestamps: false,
    createdAt: false,
    updatedAt: false,
  });
}

module.exports = getUserModel;