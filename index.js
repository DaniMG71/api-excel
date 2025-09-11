require('dotenv').config();

const express = require('express');
const cors = require('cors');
const sequelize = require('./config/database');
const getTicketModel = require('./models/DynamicTicket');

const app = express();
app.use(express.json());
app.use(cors());

// Verificar conexión
sequelize.authenticate()
  .then(() => console.log('✅ Conectado a SQL Server'))
  .catch(err => console.error('❌ Error al conectar a DB:', err));

// Inicializar modelo dinámico
let Ticket;

getTicketModel().then(model => {
  Ticket = model;

  // GET: todos los tickets
  app.get('/tickets', async (req, res) => {
    try {
      const tickets = await Ticket.findAll();
      res.json(tickets);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Error obteniendo tickets' });
    }
  });

  // GET: ticket por ID
  app.get('/tickets/:id', async (req, res) => {
    try {
      const ticket = await Ticket.findByPk(req.params.id);
      if (!ticket) return res.status(404).json({ error: 'Ticket no encontrado' });
      res.json(ticket);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Error obteniendo ticket' });
    }
  });

  // POST: crear ticket
  app.post('/tickets', async (req, res) => {
    try {
      const nuevoTicket = await Ticket.create(req.body);
      res.json(nuevoTicket);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Error creando ticket' });
    }
  });

  // DELETE: eliminar ticket
  app.delete('/tickets/:id', async (req, res) => {
    try {
      const eliminado = await Ticket.destroy({ where: { NUMERO_DE_TICKET: req.params.id } });
      if (!eliminado) return res.status(404).json({ error: 'Ticket no encontrado' });
      res.json({ message: 'Ticket eliminado correctamente' });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Error eliminando ticket' });
    }
  });

  // Iniciar servidor
  const PORT = 3001;
  app.listen(PORT, () => console.log(`Servidor corriendo en http://localhost:${PORT}`));
});
