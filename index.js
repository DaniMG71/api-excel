require('dotenv').config();
const express = require('express');
const cors = require('cors');
const sequelize = require('./config/database');
const getTicketModel = require('./models/DynamicTicket');
const getTiendaModel = require('./models/DynamicTienda');
const ldap = require('ldapjs');

const app = express();
app.use(express.json());
app.use(cors());

// Configuración LDAP
const { LDAP_URL, LDAP_BASE_DN, LDAP_DOMAIN } = process.env;

// ==========================
// FUNCIONES DE UTILIDAD
// ==========================

// Convertir fecha DD/MM/YYYY a YYYY-MM-DD
function parseFecha(fechaStr) {
  if (!fechaStr) return null;
  const partes = fechaStr.split('/');
  if (partes.length !== 3) return null;
  const [dd, mm, yyyy] = partes.map(p => parseInt(p, 10));
  if (isNaN(dd) || isNaN(mm) || isNaN(yyyy)) return null;
  const fechaISO = new Date(yyyy, mm - 1, dd);
  if (isNaN(fechaISO.getTime())) return null;
  return `${fechaISO.getFullYear()}-${String(fechaISO.getMonth() + 1).padStart(2, '0')}-${String(fechaISO.getDate()).padStart(2, '0')}`;
}

// Convertir hora HH:mm:ss (AM/PM) a formato 24h
function parseHora(horaStr) {
  if (!horaStr) return null;
  const match = horaStr.match(/^(\d{1,2}):(\d{2}):(\d{2})(?:\s?(AM|PM))?$/i);
  if (!match) return null;
  let [_, hh, mm, ss, ampm] = match;
  hh = parseInt(hh, 10);
  if (ampm) {
    if (ampm.toUpperCase() === 'PM' && hh < 12) hh += 12;
    if (ampm.toUpperCase() === 'AM' && hh === 12) hh = 0;
  }
  return `${String(hh).padStart(2, '0')}:${mm}:${ss}`;
}

// Normalizar datos de ticket
function normalizeTicketData(data) {
  if (data.año !== undefined) {
    data.anio = data.año;
    delete data.año;
  }

  const fechas = [
    'fecha_novedad',
    'fecha_resolucion',
    'fecha_reporte',
    'fecha_inicio_plan_accion',
    'fecha_fin_plan_accion',
  ];
  fechas.forEach(f => { if (data[f]) data[f] = parseFecha(data[f]); });

  const horas = [
    'hora_falla',
    'hora_reporte',
    'hora_cierre_im',
    'hora_inicio_ventana',
    'hora_fin_ventana',
  ];
  horas.forEach(h => { if (data[h]) data[h] = parseHora(data[h]); });

  const boolFields = [
    'aplica_kpi',
    'afecta_venta_pos',
    'afecta_sap',
    'afecta_genesix',
    'afecta_vtex',
    'reunion_post_incidente',
  ];
  boolFields.forEach(b => {
    if (data[b] !== undefined) {
      const val = String(data[b]).toLowerCase();
      data[b] = val === 'sí' || val === 'si' || val === 'true' || val === '1';
    }
  });

  const intFields = [
    'numero_ticket',
    'mes',
    'anio',
    'tiempo_indisp_total',
    'tiempo_indisp_ventana',
    'tiempo_notificacion',
  ];
  intFields.forEach(n => {
    if (data[n] !== undefined) {
      const num = Number(data[n]);
      data[n] = isNaN(num) ? null : num;
    }
  });

  return data;
}

// ==========================
// FUNCIÓN PRINCIPAL
// ==========================

async function startServer() {
  try {
    // Conexión a base de datos
    await sequelize.authenticate();
    console.log('✅ Conectado a PostgreSQL');

    // Cargar modelos dinámicos al inicio
    const [Ticket, Tienda] = await Promise.all([
      getTicketModel(),
      getTiendaModel(),
    ]);
    console.log('📦 Modelos cargados correctamente');

    // ======================
    // ENDPOINTS DE TICKETS
    // ======================
    app.get('/tickets', async (req, res) => {
      try {
        const tickets = await Ticket.findAll();
        res.json(tickets);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error obteniendo tickets' });
      }
    });

    app.get('/tickets/:id', async (req, res) => {
      try {
        const ticket = await Ticket.findByPk(req.params.id);
        if (!ticket) return res.status(404).json({ error: 'Ticket no encontrado' });
        res.json(ticket);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error obteniendo ticket' });
      }
    });

    app.post('/tickets', async (req, res) => {
      try {
        const data = normalizeTicketData(req.body);
        const nuevo = await Ticket.create(data);
        res.json(nuevo);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error creando ticket' });
      }
    });

    app.put('/tickets/:id', async (req, res) => {
      try {
        const { id } = req.params;
        const data = normalizeTicketData(req.body);
        const [updated] = await Ticket.update(data, { where: { numero_ticket: id } });
        if (!updated) return res.status(404).json({ error: 'Ticket no encontrado' });
        const ticket = await Ticket.findOne({ where: { numero_ticket: id } });
        res.json(ticket);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error actualizando ticket' });
      }
    });

    app.delete('/tickets/:id', async (req, res) => {
      try {
        const deleted = await Ticket.destroy({ where: { numero_ticket: req.params.id } });
        if (!deleted) return res.status(404).json({ error: 'Ticket no encontrado' });
        res.json({ message: 'Ticket eliminado correctamente' });
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error eliminando ticket' });
      }
    });

    // ======================
    // ENDPOINTS DE TIENDAS
    // ======================
    app.get('/tiendas', async (req, res) => {
      try {
        const tiendas = await Tienda.findAll();
        res.json(tiendas);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error obteniendo tiendas' });
      }
    });

    app.get('/tiendas/:cod_sap', async (req, res) => {
      try {
        const tienda = await Tienda.findByPk(req.params.cod_sap);
        if (!tienda) return res.status(404).json({ error: 'Tienda no encontrada' });
        res.json(tienda);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error obteniendo tienda' });
      }
    });

    app.post('/tiendas', async (req, res) => {
      try {
        const nueva = await Tienda.create(req.body);
        res.status(201).json(nueva);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error creando tienda' });
      }
    });

    app.put('/tiendas/:cod_sap', async (req, res) => {
      try {
        const { cod_sap } = req.params;
        const [updated] = await Tienda.update(req.body, { where: { cod_sap } });
        if (!updated) return res.status(404).json({ error: 'Tienda no encontrada' });
        const tienda = await Tienda.findOne({ where: { cod_sap } });
        res.json(tienda);
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error actualizando tienda' });
      }
    });

    app.delete('/tiendas/:cod_sap', async (req, res) => {
      try {
        const deleted = await Tienda.destroy({ where: { cod_sap: req.params.cod_sap } });
        if (!deleted) return res.status(404).json({ error: 'Tienda no encontrada' });
        res.json({ message: 'Tienda eliminada correctamente' });
      } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Error eliminando tienda' });
      }
    });

    // ======================
    // LOGIN LDAP
    // ======================
    app.post('/auth', (req, res) => {
      const { user, pass } = req.body;
      if (!user || !pass) return res.status(400).json({ error: 'Usuario y contraseña requeridos' });

      const client = ldap.createClient({ url: LDAP_URL });
      const ldapUser = `${user}@${LDAP_DOMAIN}`;

      client.bind(ldapUser, pass, (err) => {
        if (err) {
          console.error('❌ Credenciales inválidas:', err.message);
          client.unbind();
          return res.status(401).json({ error: 'Credenciales inválidas' });
        }

        const opts = {
          filter: `(sAMAccountName=${user})`,
          scope: 'sub',
          attributes: ['sAMAccountName', 'displayName', 'memberOf'],
        };

        client.search(LDAP_BASE_DN, opts, (err, searchRes) => {
          if (err) {
            client.unbind();
            return res.status(500).json({ error: 'Error en búsqueda LDAP' });
          }

          let userEntry = null;
          searchRes.on('searchEntry', (entry) => (userEntry = entry.object));
          searchRes.on('end', () => {
            client.unbind();
            if (!userEntry) return res.status(404).json({ error: 'Usuario no encontrado' });
            res.json({ success: true, user: userEntry });
          });
        });
      });
    });


    // ======================
    // ENDPOINT PARA COLUMNAS
    // ======================
    app.get('/columns', async (req, res) => {
      try {
        const [result] = await sequelize.query(`
      SELECT column_name 
      FROM information_schema.columns 
      WHERE table_name = 'tickets'
      ORDER BY ordinal_position;
    `);

        const columnas = result.map(r => r.column_name);
        res.json({ success: true, columns: columnas });
      } catch (err) {
        console.error('❌ Error obteniendo columnas:', err);
        res.status(500).json({ success: false, error: 'Error obteniendo columnas' });
      }
    });

    // ======================
    // INICIAR SERVIDOR
    // ======================
    const PORT = process.env.PORT || 3001;
    app.listen(PORT, () => console.log(`🚀 Servidor corriendo en http://localhost:${PORT}`));

  } catch (err) {
    console.error('❌ Error al iniciar servidor:', err);
    process.exit(1);
  }
}

startServer();
