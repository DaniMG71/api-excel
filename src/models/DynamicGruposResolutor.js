// src/models/DynamicGruposResolutor.js
const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

async function getGruposResolutorModel() {
  // Traer información de las columnas de la tabla
  const [columns] = await sequelize.query(`
    SELECT column_name AS "COLUMN_NAME", data_type AS "DATA_TYPE"
    FROM information_schema.columns
    WHERE table_name = 'grupos_resolutor'
      AND table_schema = 'public'
  `);

  if (!columns.length) throw new Error('No se encontraron columnas en GruposResolutor');

  const modelDefinition = {};

  // Mapear tipos PostgreSQL a Sequelize
  const sqlToSequelize = {
    bigint: DataTypes.BIGINT,
    integer: DataTypes.INTEGER,
    "character varying": DataTypes.STRING,
    text: DataTypes.TEXT,
    date: DataTypes.DATEONLY,
    time: DataTypes.TIME,
    "timestamp without time zone": DataTypes.DATE,
    "timestamp with time zone": DataTypes.DATE,
    boolean: DataTypes.BOOLEAN,
    bit: DataTypes.BOOLEAN,
    numeric: DataTypes.DECIMAL,
    "double precision": DataTypes.FLOAT,
  };

  for (const col of columns) {
    const name = col.COLUMN_NAME;
    const type = sqlToSequelize[col.DATA_TYPE.toLowerCase()] || DataTypes.STRING;

    modelDefinition[name] = { type };

    // Detectar primary key
    if (name === 'id_grupo') {
      modelDefinition[name].primaryKey = true;
      modelDefinition[name].autoIncrement = true;
    }
  }

  return sequelize.define('GrupoResolutor', modelDefinition, {
    tableName: 'grupos_resolutor',
    timestamps: false,  // DESHABILITAR timestamps automáticos para evitar conflictos
    createdAt: false,   // No mapear
    updatedAt: false,   // No mapear
  });
}

module.exports = getGruposResolutorModel;