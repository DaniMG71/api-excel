const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

async function getDynamicReunionesAsistentesModel() {
  const [columns] = await sequelize.query(`
    SELECT column_name AS "COLUMN_NAME", data_type AS "DATA_TYPE"
    FROM information_schema.columns
    WHERE table_name = 'reunionesasistentes'
      AND table_schema = 'public';
  `);

  if (!columns.length) throw new Error('No se encontraron columnas en reunionesasistentes');

  const modelDefinition = {};

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
    const name = col.COLUMN_NAME.toLowerCase();
    const type = sqlToSequelize[col.DATA_TYPE.toLowerCase()] || DataTypes.STRING;

    modelDefinition[name] = { type };

    // 🔑 Clave primaria compuesta: id_reunion e id_personal
    if (name === 'id_reunion' || name === 'id_personal') {
      modelDefinition[name].primaryKey = true;  // Ambas son parte de la PK
      modelDefinition[name].allowNull = false;
      // No autoIncrement para ninguna, ya que es FK
    }

    // 🔗 FK a reuniones
    if (name === 'id_reunion') {
      modelDefinition[name].references = {
        model: 'reuniones',  // Tabla relacionada
        key: 'id_reunion',
      };
      modelDefinition[name].onDelete = 'CASCADE';
    }

    // 🔗 FK a personal
    if (name === 'id_personal') {
      modelDefinition[name].references = {
        model: 'personal',
        key: 'id_personal',  // Corregido: la columna en personal es id_personal
      };
      modelDefinition[name].onDelete = 'CASCADE';
    }
  }

  // 🔧 FORZAR LA INCLUSIÓN DE 'asistio' SI NO EXISTE EN LA TABLA
  if (!modelDefinition.asistio) {
    modelDefinition.asistio = {
      type: DataTypes.BOOLEAN,
      defaultValue: false,
      allowNull: false,
    };
    console.log('[INFO] Campo "asistio" forzado en el modelo ReunionesAsistentes');
  }

  return sequelize.define('ReunionesAsistentes', modelDefinition, {
    tableName: 'reunionesasistentes',
    timestamps: false,
  });
}

module.exports = getDynamicReunionesAsistentesModel;