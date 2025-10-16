const { DataTypes } = require("sequelize");
const sequelize = require("../config/database");

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
    const type = /* tu lógica para mapear tipos */;

    modelDefinition[name] = { type };

    if (name.toLowerCase() === "id_asistente") {
      modelDefinition[name].primaryKey = true;
      modelDefinition[name].autoIncrement = true;  // Asegúrate de que esto esté activado
      modelDefinition[name].allowNull = false;
    }
  }

  return sequelize.define("Asistentes", modelDefinition, {
    tableName: "asistentes",
    timestamps: false,
  });
}

module.exports = getAsistenteModel;