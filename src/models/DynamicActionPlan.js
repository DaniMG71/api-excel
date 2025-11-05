// models/DynamicActionPlan.js
const { DataTypes } = require("sequelize");
const sequelize = require('../config/database');

async function getActionPlanModel() {
    // 1️⃣ Obtener columnas desde la tabla
    const [columns] = await sequelize.query(`
    SELECT column_name AS "COLUMN_NAME", data_type AS "DATA_TYPE"
    FROM information_schema.columns
    WHERE LOWER(table_name) = 'planaccion'
      AND table_schema = 'public';
  `);

    if (!columns.length) throw new Error("No se encontraron columnas en Plan de Acción");

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
        json: DataTypes.JSON,
        "character varying[]": DataTypes.ARRAY(DataTypes.STRING),
    };

    for (const col of columns) {
        const name = col.COLUMN_NAME;
        const type =
            sqlToSequelize[col.DATA_TYPE.toLowerCase()] || DataTypes.STRING;

        modelDefinition[name] = { type };

        // Detectar primary key
        if (name.toLowerCase() === "id_plan_accion") {
            modelDefinition[name].primaryKey = true;
            modelDefinition[name].autoIncrement = true;
            modelDefinition[name].allowNull = false;
        }
    }

    if (!modelDefinition.novedad) {
        modelDefinition.novedad = {
            type: DataTypes.TEXT,  // O DataTypes.STRING si prefieres VARCHAR
            allowNull: true,
            defaultValue: "",
        };
    }   

     modelDefinition.personal = {
        type: DataTypes.JSON,
        allowNull: true,
        defaultValue: [],
    };
    return sequelize.define("PlanAccion", modelDefinition, {
        tableName: "planaccion",
        timestamps: false,
        createdAt: false,
        updatedAt: false,
    });
}

module.exports = getActionPlanModel;