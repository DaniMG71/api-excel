const { DataTypes } = require("sequelize");
const sequelize = require('../config/database');

async function getAsistenteModel() {
  const [columns] = await sequelize.query(`
    SELECT column_name AS "COLUMN_NAME", data_type AS "DATA_TYPE"
    FROM information_schema.columns
    WHERE LOWER(table_name) = 'asistentes'
      AND table_schema = 'public';
  `);

  if (!columns.length) throw new Error("No se encontraron columnas en Asistentes");

  const modelDefinition = {};

  for (const col of columns) {
    const name = col.COLUMN_NAME;
    const dataType = col.DATA_TYPE.toLowerCase();

    // 🔧 Lógica para mapear tipos PostgreSQL → Sequelize
    let type;
    if (dataType.includes("character") || dataType.includes("text")) type = DataTypes.STRING;
    else if (dataType.includes("integer")) type = DataTypes.INTEGER;
    else if (dataType.includes("double") || dataType.includes("numeric") || dataType.includes("decimal")) type = DataTypes.FLOAT;
    else if (dataType.includes("boolean")) type = DataTypes.BOOLEAN;
    else if (dataType.includes("date")) type = DataTypes.DATE;
    else type = DataTypes.STRING; // valor por defecto

    modelDefinition[name] = { type };

    // 🔑 Detecta clave primaria
    if (name.toLowerCase() === "id_asistente") {
      modelDefinition[name].primaryKey = true;
      modelDefinition[name].autoIncrement = true;
      modelDefinition[name].allowNull = false;
    }
  }

  return sequelize.define("Asistentes", modelDefinition, {
    tableName: "asistentes",
    timestamps: false,
  });
}

module.exports = getAsistenteModel;
