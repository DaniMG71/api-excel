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
const LDAP_URL = process.env.LDAP_URL;
const LDAP_BASE_DN = process.env.LDAP_BASE_DN ;
const LDAP_DOMAIN = process.env.LDAP_DOMAIN;

// Verificar conexión DB
sequelize.authenticate()
  .then(() => console.log('✅ Conectado a postgres'))
  .catch(err => console.error('❌ Error al conectar a DB:', err));

// Inicializar modelo dinámico
let Ticket, Tienda;

// Función para convertir fecha DD/MM/YYYY a YYYY-MM-DD
function parseFecha(fechaStr) {
  if (!fechaStr) return null;
  const partes = fechaStr.split('/');
  if (partes.length !== 3) return null;
  const [dd, mm, yyyy] = partes.map(p => parseInt(p, 10));
  if (
    isNaN(dd) || isNaN(mm) || isNaN(yyyy) ||
    dd < 1 || dd > 31 ||
    mm < 1 || mm > 12 ||
    yyyy < 1900
  ) return null;

  const fechaISO = new Date(yyyy, mm - 1, dd);
  if (isNaN(fechaISO.getTime())) return null;

  const yyyyStr = fechaISO.getFullYear();
  const mmStr = String(fechaISO.getMonth() + 1).padStart(2, '0');
  const ddStr = String(fechaISO.getDate()).padStart(2, '0');
  return `${yyyyStr}-${mmStr}-${ddStr}`;
}

// Función para convertir hora a formato HH:mm:ss
function parseHora(horaStr) {
  if (!horaStr) return null;
  const hora24 = horaStr.match(/^(\d{1,2}):(\d{2}):(\d{2})(?:\s?(AM|PM))?$/i);
  if (!hora24) return null;

  let hh = parseInt(hora24[1], 10);
  const mm = hora24[2];
  const ss = hora24[3];
  const ampm = hora24[4];

  if (ampm) {
    if (ampm.toUpperCase() === 'PM' && hh < 12) hh += 12;
    if (ampm.toUpperCase() === 'AM' && hh === 12) hh = 0;
  }

  if (hh < 0 || hh > 23) return null;

  return `${String(hh).padStart(2, '0')}:${mm}:${ss}`;
}

// Función para normalizar todo el objeto ticket antes de guardar
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
  fechas.forEach(f => {
    if (data[f]) {
      data[f] = parseFecha(data[f]);
    }
  });

  const horas = [
    'hora_falla',
    'hora_reporte',
    'hora_cierre_im',
    'hora_inicio_ventana',
    'hora_fin_ventana',
  ];
  horas.forEach(h => {
    if (data[h]) {
      data[h] = parseHora(data[h]);
    }
  });

  const boolFields = [
    'aplica_kpi',
    'afecta_venta_pos',
    'afecta_sap',
    'afecta_genesix',
    'afecta_vtex',
    'reunion_post_incidente'
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
    'tiempo_notificacion'
  ];
  intFields.forEach(n => {
    if (data[n] !== undefined) {
      const num = Number(data[n]);
      data[n] = isNaN(num) ? null : num;
    }
  });

  return data;
}

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
    const normalizedData = normalizeTicketData(req.body);
    const nuevoTicket = await Ticket.create(normalizedData);
    res.json(nuevoTicket);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Error creando ticket' });
  }
});

app.put('/tickets/:id', async (req, res) => {
  try {
    const { id } = req.params;
    console.log("✏️ ID recibido para editar:", id);

    // Normalizar datos si usas la misma función que al crear
    const updatedData = normalizeTicketData(req.body);

    const [updated] = await Ticket.update(updatedData, {
      where: { numero_ticket: id }
    });

    if (updated === 0) {
      return res.status(404).json({ error: '❌ Ticket no encontrado' });
    }

    const ticketActualizado = await Ticket.findOne({ where: { numero_ticket: id } });
    res.json({ message: '✅ Ticket actualizado correctamente', ticket: ticketActualizado });
  } catch (error) {
    console.error("🔥 Error actualizando ticket:", error);
    res.status(500).json({ error: 'Error actualizando ticket', details: error.message });
  }
});


  app.delete('/tickets/:id', async (req, res) => {
  try {
    console.log("🗑️ ID recibido para eliminar:", req.params.id)

    const eliminado = await Ticket.destroy({
      where: { numero_ticket: req.params.id } // 👈 Exactamente así
    })

    console.log("Resultado destroy:", eliminado)

    if (!eliminado) return res.status(404).json({ error: 'Ticket no encontrado' })
    res.json({ message: '✅ Ticket eliminado correctamente' })
  } catch (error) {
    console.error("🔥 Error eliminando ticket:", error)
    res.status(500).json({ error: 'Error eliminando ticket', details: error.message })
  }
})





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
  // endpoints de TIENDAS
  // ======================

  getTiendaModel().then(model => {
  Tienda = model;

  // Obtener todas las tiendas
app.get("/tiendas", async (req, res) => {
  try {
    const tiendas = await Tienda.findAll();
    //console.log("📊 Resultado de Tienda.findAll():", tiendas);
    res.json(tiendas);
  } catch (error) {
    console.error("❌ Error obteniendo tiendas:", error);
    res.status(500).json({ error: "Error al obtener tiendas" });
  }
});

// Obtener tienda por COD SAP (usa findByPk ya que es PK)
app.get("/tiendas/:cod_sap", async (req, res) => {
  try {
    const { cod_sap } = req.params;
    console.log(`🔍 Buscando tienda por COD SAP: ${cod_sap}`);  // DEBUG

    const tienda = await Tienda.findByPk(cod_sap);  // Usa PK directamente (eficiente)
    if (!tienda) {
      return res.status(404).json({ error: "Tienda no encontrada" });
    }
    console.log(`✅ Tienda encontrada: ${tienda.get('NOMBRE_PTO_OPERACIONAL')}`);  // DEBUG
    res.json(tienda);
  } catch (error) {
    console.error("❌ Error obteniendo tienda:", error);
    res.status(500).json({ error: "Error al obtener tienda" });
  }
});

app.delete("/tiendas/:cod_sap", async (req, res) => {
  const { cod_sap } = req.params;
  try {
    
    // Sequelize: eliminar por PK
    const eliminado = await Tienda.destroy({
      where: { cod_sap }
    });

    if (eliminado === 0) {
      return res.status(404).json({ error: "Tienda no encontrada" });
    }

    console.log("✅ Tienda eliminada correctamente");
    res.status(200).json({ message: "Tienda eliminada correctamente ✅" });
  } catch (err) {
    console.error("❌ Error en DELETE /tiendas/:cod_sap:", err.message);
    res.status(500).json({ error: err.message });
  }
});

// ✅ Actualizar tienda por COD_SAP
app.put("/tiendas/:cod_sap", async (req, res) => {
  try {
    const cod_sap_param = req.params.cod_sap; // 🔹 usar un nombre diferente al de la columna
    console.log("🛠️ Actualizando tienda:", cod_sap_param);

    const [updated] = await Tienda.update(req.body, {
      where: { cod_sap: cod_sap_param } // 🔹 coincide con la columna en mayúsculas
    });

    if (updated === 0) {
      return res.status(404).json({ error: "❌ Tienda no encontrada" });
    }

    const tiendaActualizada = await Tienda.findOne({ where: { cod_sap: cod_sap_param } });
    res.json(tiendaActualizada);
  } catch (error) {
    console.error("❌ Error al actualizar tienda:", error);
    res.status(500).json({ error: "Error al actualizar tienda" });
  }
});

// Crear tienda (ejemplo, si lo usas – usa nombres exactos)
app.post("/tiendas", async (req, res) => {
  try {
    const nuevaTienda = await Tienda.create(req.body);  // req.body debe tener 'COD_SAP', etc.
    res.status(201).json(nuevaTienda);
  } catch (error) {
    console.error("❌ Error creando tienda:", error);
    res.status(500).json({ error: "Error al crear tienda" });
  }
});
});

  // ======================
  // INICIAR SERVIDOR
  // ======================
  const PORT = process.env.PORT || 3001;
  app.listen(PORT, () => console.log(`🚀 Servidor corriendo en http://localhost:${PORT}`));
});
