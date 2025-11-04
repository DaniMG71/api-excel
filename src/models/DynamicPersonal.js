const { DataTypes } = require("sequelize");
const sequelize = require('../config/database');

async function getPersonalModel() {
  const [columns] = await sequelize.query(`
    SELECT column_name AS "COLUMN_NAME", data_type AS "DATA_TYPE"
    FROM information_schema.columns
    WHERE LOWER(table_name) = 'personal'
      AND table_schema = 'public';
  `);

  if (!columns.length) throw new Error("No se encontraron columnas en Personal");

  const modelDefinition = {};

  for (const col of columns) {
    const name = col.COLUMN_NAME;
    const dataType = col.DATA_TYPE.toLowerCase();

    // 🔧 Mapear tipos PostgreSQL → Sequelize (igual que antes)
    let type;
    if (dataType.includes("character") || dataType.includes("text")) type = DataTypes.STRING;
    else if (dataType.includes("integer")) type = DataTypes.INTEGER;
    else if (dataType.includes("bigint")) type = DataTypes.BIGINT;
    else if (dataType.includes("double") || dataType.includes("numeric") || dataType.includes("decimal")) type = DataTypes.FLOAT;
    else if (dataType.includes("boolean")) type = DataTypes.BOOLEAN;
    else if (dataType.includes("date") || dataType.includes("timestamp")) type = DataTypes.DATE;
    else type = DataTypes.STRING;

    // ⚠️ Saltar createdAt y updatedAt
    if (name === 'createdAt' || name === 'updatedAt') continue;

    modelDefinition[name] = { type };

    // 🔑 Configurar clave primaria como INTEGER auto-increment
    if (name.toLowerCase() === "id_personal") {
      modelDefinition[name].primaryKey = true;
      modelDefinition[name].allowNull = false;
      modelDefinition[name].autoIncrement = true;  // ✅ Auto-increment para INTEGER
      // Quitar defaultValue si estaba (no aplica a INTEGER)
    }
  }

  return sequelize.define("Personal", modelDefinition, {
    tableName: "personal",
    timestamps: true,
    createdAt: "createdAt",
    updatedAt: "updatedAt",
  });
}

module.exports = getPersonalModel;
