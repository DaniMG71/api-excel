require('dotenv').config();

const express = require('express');
const cors = require('cors');
const sequelize = require('./config/database');
const getTicketModel = require('./models/DynamicTicket');
const ldap = require('ldapjs');

const app = express();
app.use(express.json());
app.use(cors());

// Configuración LDAP
const LDAP_URL = process.env.LDAP_URL;
const LDAP_BASE_DN = process.env.LDAP_BASE_DN ;
const LDAP_DOMAIN = process.env.LDAP_DOMAIN;

// Verificar conexión DB
sequelize.authenticate()
  .then(() => console.log('✅ Conectado a SQL Server'))
  .catch(err => console.error('❌ Error al conectar a DB:', err));

// Inicializar modelo dinámico
let Ticket;

getTicketModel().then(model => {
  Ticket = model;

  // ======================
  // ENDPOINTS DE TICKETS
  // ======================
  app.get('/tickets', async (req, res) => {
    try {
      const tickets = await Ticket.findAll();
      res.json(tickets);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Error obteniendo tickets' });
    }
  });

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

  app.post('/tickets', async (req, res) => {
    try {
      const nuevoTicket = await Ticket.create(req.body);
      res.json(nuevoTicket);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Error creando ticket' });
    }
  });

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

  // ======================
  // ENDPOINT DE LOGIN LDAP
  // ======================
  app.post('/auth', (req, res) => {
    const { user, pass } = req.body;

    if (!user || !pass) {
      return res.status(400).json({ error: "Usuario y contraseña requeridos" });
    }

    const client = ldap.createClient({ url: LDAP_URL });
    const ldapUser = `${user}@${LDAP_DOMAIN}`;

    client.bind(ldapUser, pass, (err) => {
      if (err) {
        console.error("❌ Credenciales inválidas:", err.message);
        client.unbind();
        return res.status(401).json({ error: "Credenciales inválidas" });
      }

      const opts = {
        filter: `(sAMAccountName=${user})`,
        scope: "sub",
        attributes: ["sAMAccountName", "displayName", "memberOf"],
      };

      client.search(LDAP_BASE_DN, opts, (err, searchRes) => {
        if (err) {
          console.error("❌ Error en búsqueda LDAP:", err);
          client.unbind();
          return res.status(500).json({ error: "Error en búsqueda LDAP" });
        }

        let userEntry = null;

        searchRes.on("searchEntry", (entry) => {
          userEntry = entry.object;
        });

        searchRes.on("end", () => {
          client.unbind();
          if (!userEntry) {
            return res.status(404).json({ error: "Usuario no encontrado" });
          }
          return res.json({
            success: true,
            user: userEntry
          });
        });
      });
    });
  });

  // ======================
  // INICIAR SERVIDOR
  // ======================
  const PORT = process.env.PORT || 3001;
  app.listen(PORT, () => console.log(`🚀 Servidor corriendo en http://localhost:${PORT}`));
});
