// models/TicketTienda.js
const { DataTypes } = require("sequelize");
const sequelize = require('../config/database');

const TicketTienda = sequelize.define("TicketTienda", {
  id: {
    type: DataTypes.INTEGER,
    autoIncrement: true,
    primaryKey: true,
  },
  numero_ticket: {
    type: DataTypes.BIGINT,
    allowNull: false,
  },
  cod_sap: {
    type: DataTypes.STRING,
    allowNull: false,
  }
}, {
  tableName: "ticket_tienda",
  timestamps: false,
  indexes: [
    { fields: ['numero_ticket'] },  // Index para FK
    { fields: ['cod_sap'] },        // Index para FK
    { unique: true, fields: ['numero_ticket', 'cod_sap'] }  // Evita duplicados ticket-tienda
  ],
});

module.exports = TicketTienda;