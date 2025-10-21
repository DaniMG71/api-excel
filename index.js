require('dotenv').config();
const express = require('express');
const cors = require('cors');
const sequelize = require('./config/database');
const getTicketModel = require('./models/DynamicTicket');
const getTiendaModel = require('./models/DynamicTienda');
const getActionPlanModel = require('./models/DynamicActionPlan');
const getMeetingModel = require('./models/DynamicMeeting');
const getAsistenteModel = require('./models/DynamicAsistentes');
const getDynamicReunionesAsistentesModel = require('./models/DynamicReunionesAsistentes');
const TicketTienda = require('./models/TicketTienda');
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
    const [Ticket, Tienda, PlanAccion, Reunion, Asistente, ReunionAsistente] = await Promise.all([
      getTicketModel(),
      getTiendaModel(),
      getActionPlanModel(),
      getMeetingModel(),
      getAsistenteModel(),
      getDynamicReunionesAsistentesModel(),
      
    ]);
    console.log('📦 Modelos cargados correctamente');
// Relaciones
Ticket.hasMany(PlanAccion, { foreignKey: "numero_ticket", as: "planes" });
PlanAccion.belongsTo(Ticket, { foreignKey: "numero_ticket" });

PlanAccion.hasMany(Reunion, { foreignKey: "id_plan_accion", as: "reuniones" });
Reunion.belongsTo(PlanAccion, { foreignKey: "id_plan_accion" });

// belongsToMany: muchos a muchos entre reuniones y asistentes
Reunion.belongsToMany(Asistente, {
  through: "reunionesasistentes",  // nombre exacto de tu tabla intermedia
  foreignKey: "id_reunion",         // en la tabla intermedia
  otherKey: "id_asistente",         // en la tabla intermedia
  as: "asistentes",
});

Asistente.belongsToMany(Reunion, {
  through: "reunionesasistentes",
  foreignKey: "id_asistente",
  otherKey: "id_reunion",
  as: "reuniones",
});

// Sync tablas (solo para dev - quítalo en prod para evitar alteraciones)
sequelize.sync({ alter: true }).then(() => console.log('🔄 Tablas sincronizadas'));



    // ======================
// ENDPOINTS DE TICKETS
// ======================
app.get('/tickets', async (req, res) => {
  try {
    console.log('🔍 Obteniendo todos los tickets con tiendas asociadas...');
    
    // Obtener tickets principales
    const tickets = await Ticket.findAll();
    
    // Para cada ticket, obtener TODAS las tiendas asociadas (primaria + adicionales)
    const ticketsConTiendas = await Promise.all(tickets.map(async (ticket) => {
      // 1. Obtener relaciones de ticket_tienda
      const relaciones = await TicketTienda.findAll({ 
        where: { numero_ticket: ticket.numero_ticket } 
      });
      
      // 2. Obtener cod_saps únicos
      const codSaps = [...new Set(relaciones.map(r => r.cod_sap))];  // Evita duplicados
      
      console.log(`   - Ticket ${ticket.numero_ticket}: ${codSaps.length} tiendas asociadas (${codSaps.join(', ')})`);
      
      // 3. Obtener detalles de cada tienda
      const tiendasAsociadas = await Promise.all(
        codSaps.map(cod_sap => Tienda.findByPk(cod_sap))
      );
      
      // Filtrar nulls (si alguna tienda no existe)
      const tiendasValidas = tiendasAsociadas.filter(t => t !== null);
      
      // Agregar al ticket (mantén los campos originales como primaria)
      return {
        ...ticket.dataValues,  // Todos los campos del ticket
        tiendasAsociadas: tiendasValidas,  // Array de todas las tiendas
        numTiendas: tiendasValidas.length  // Útil para UI
      };
    }));

    console.log(`✅ Enviando ${ticketsConTiendas.length} tickets con tiendas asociadas`);
    res.json(ticketsConTiendas);
  } catch (error) {
    console.error('❌ Error obteniendo tickets con tiendas:', error);
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
  const t = await sequelize.transaction();  // 🧠 Transacción para rollback si falla
  try {
    console.log('📝 Creando nuevo ticket con tiendas...');
    const data = normalizeTicketData(req.body);
    
    // 🔥 OBTENER ARRAY DE TIENDAS DEL BODY (ajusta el nombre si tu frontend envía diferente, ej. 'tiendasSeleccionadas')
    const codigosTiendas = data.codigosTiendas || [];  // Ej: ["M026", "M028"]
    if (codigosTiendas.length === 0) {
      throw new Error('Debe seleccionar al menos una tienda');
    }
    
    console.log(`   - Tiendas seleccionadas: ${codigosTiendas.join(', ')} (${codigosTiendas.length} total)`);
    
    // 1️⃣ OBTENER DETALLES DE LAS TIENDAS
    const tiendas = await Promise.all(
      codigosTiendas.map(cod_sap => Tienda.findByPk(cod_sap, { transaction: t }))
    );
    
    // Filtrar tiendas inválidas (si alguna no existe)
    const tiendasValidas = tiendas.filter(tienda => tienda !== null);
    if (tiendasValidas.length === 0) {
      throw new Error('Ninguna tienda seleccionada existe en la base de datos');
    }
    
    // 2️⃣ LLENAR CAMPOS PRIMARIOS CON LA PRIMERA TIENDA (para compatibilidad)
    const primeraTienda = tiendasValidas[0];
    data.codigo_tienda = primeraTienda.cod_sap;
    data.tienda = primeraTienda.nombre_pto_operacional;
    data.unidad_negocio = primeraTienda.uunn;
    data.bandera = primeraTienda.bandera;
    data.region = primeraTienda.region;
    
    console.log(`   - Primaria: ${data.codigo_tienda} (${data.tienda})`);
    
    // 3️⃣ CREAR EL TICKET PRINCIPAL (sin codigo_tienda en el create, ya que lo agregamos arriba)
    const nuevoTicket = await Ticket.create(data, { transaction: t });
    
    // 4️⃣ CREAR RELACIONES EN TICKET_TIENDA PARA TODAS LAS TIENDAS
    for (const tienda of tiendasValidas) {
      await TicketTienda.create({
        numero_ticket: nuevoTicket.numero_ticket,
        cod_sap: tienda.cod_sap
      }, { transaction: t });
      
      console.log(`   - Relación creada: Ticket ${nuevoTicket.numero_ticket} ↔ Tienda ${tienda.cod_sap}`);
    }
    
    // 5️⃣ Opcional: Limpiar el array del body para no guardarlo en el ticket
    delete data.codigosTiendas;
    
    // 6️⃣ COMMIT Y RESPUESTA
    await t.commit();
    
    // Fetch el ticket completo con tiendas para devolver (usa el GET mejorado si lo tienes)
    const ticketCompleto = await Ticket.findByPk(nuevoTicket.numero_ticket, {
      include: []  // Si tienes asociaciones, agrégalas aquí
    });
    
    console.log(`✅ Ticket ${nuevoTicket.numero_ticket} creado con ${tiendasValidas.length} tiendas`);
    res.status(201).json(ticketCompleto);
    
  } catch (error) {
    await t.rollback();
    console.error('❌ Error creando ticket:', error);
    res.status(500).json({ error: error.message || 'Error creando ticket' });
  }
});

    // index.js (actualizado) - Solo el endpoint PUT /tickets se modifica; el resto permanece igual
// ... (código anterior sin cambios hasta app.put('/tickets/:id', ...))

app.put('/tickets/:id', async (req, res) => {
  const t = await sequelize.transaction();
  try {
    const { id } = req.params;
    let data = normalizeTicketData(req.body);
    
    // 🔥 Manejar múltiples tiendas si se envían
    const codigosTiendas = data.codigosTiendas || [];
    delete data.codigosTiendas;
    
    let syncTiendas = false;
    if (codigosTiendas.length > 0) {
      syncTiendas = true;
      
      // Obtener detalles de las tiendas
      const tiendas = await Promise.all(
        codigosTiendas.map(cod_sap => Tienda.findByPk(cod_sap, { transaction: t }))
      );
      
      // Filtrar tiendas válidas
      const tiendasValidas = tiendas.filter(tienda => tienda !== null);
      if (tiendasValidas.length === 0) {
        throw new Error('Ninguna tienda seleccionada existe en la base de datos');
      }
      if (tiendasValidas.length !== codigosTiendas.length) {
        throw new Error('Algunas tiendas no tienen código SAP asignado');
      }
      
      console.log(`   - Actualizando ticket ${id} con ${tiendasValidas.length} tiendas: ${codigosTiendas.join(', ')}`);
      
      // Establecer campos primarios desde la primera tienda (para compatibilidad)
      const primeraTienda = tiendasValidas[0];
      data.codigo_tienda = primeraTienda.cod_sap;
      data.tienda = primeraTienda.nombre_pto_operacional;
      data.unidad_negocio = primeraTienda.uunn;
      data.bandera = primeraTienda.bandera;
      data.region = primeraTienda.region;
      
      console.log(`   - Primaria actualizada: ${data.codigo_tienda} (${data.tienda})`);
      
      // Eliminar asociaciones existentes
      await TicketTienda.destroy({
        where: { numero_ticket: id },
        transaction: t
      });
      
      // Crear nuevas asociaciones
      for (const tienda of tiendasValidas) {
        await TicketTienda.create({
          numero_ticket: id,
          cod_sap: tienda.cod_sap
        }, { transaction: t });
        
        console.log(`   - Relación actualizada: Ticket ${id} ↔ Tienda ${tienda.cod_sap}`);
      }
    } else {
      console.log(`   - Actualizando ticket ${id} sin cambios en tiendas`);
    }
    
    // Actualizar el ticket principal
    const [updated] = await Ticket.update(data, {
      where: { numero_ticket: id },
      transaction: t
    });
    
    if (!updated) {
      await t.rollback();
      return res.status(404).json({ error: 'Ticket no encontrado' });
    }
    
    // Commit de la transacción
    await t.commit();
    
    // Fetch el ticket actualizado con tiendas asociadas (similar al GET /tickets)
    const ticket = await Ticket.findByPk(id);
    if (!ticket) {
      return res.status(404).json({ error: 'Ticket no encontrado después de actualización' });
    }
    
    // Obtener relaciones de ticket_tienda
    const relaciones = await TicketTienda.findAll({ 
      where: { numero_ticket: id } 
    });
    
    // Obtener cod_saps únicos
    const codSaps = [...new Set(relaciones.map(r => r.cod_sap))];
    
    // Obtener detalles de cada tienda
    const tiendasAsociadas = await Promise.all(
      codSaps.map(cod_sap => Tienda.findByPk(cod_sap))
    );
    
    // Filtrar nulls
    const tiendasValidas = tiendasAsociadas.filter(t => t !== null);
    
    // Retornar ticket con tiendas asociadas
    const ticketCompleto = {
      ...ticket.dataValues,
      tiendasAsociadas: tiendasValidas,
      numTiendas: tiendasValidas.length
    };
    
    console.log(`✅ Ticket ${id} actualizado con éxito`);
    res.json(ticketCompleto);
    
  } catch (error) {
    await t.rollback();
    console.error('❌ Error actualizando ticket:', error);
    res.status(500).json({ error: error.message || 'Error actualizando ticket' });
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
    // ENDPOINT PARA PLAN DE ACCIÓN + REUNIONES
    // ======================

app.get("/plan-accion", async (req, res) => {
  try {
    const tickets = await Ticket.findAll({
      include: [
        {
          model: PlanAccion,
          as: "planes",
          include: [
            {
              model: Reunion,
              as: "reuniones",
              include: [
                {
                  model: Asistente,
                  as: "asistentes",
                },
              ],
            },
          ],
        },
      ],
    });

    const data = tickets.map(ticket => ({
      "NUMERO DE TICKET": ticket.numero_ticket,
      "ESTADO DEL PLAN DE ACCION": ticket.estado_plan_accion || "Pendiente",
      "TIPO DE TICKET": ticket.tipo_ticket,
      planes: ticket.planes.map(plan => ({
        id_plan_accion: plan.id_plan_accion,
        numero_ticket: plan.numero_ticket,
        tipo_ticket: plan.tipo_ticket,
        estado_plan_accion: plan.estado_plan_accion,
        plan_accion: plan.plan_accion,
        servicio: plan.servicio,
        fecha_apertura: plan.fecha_apertura,
        fecha_cierre: plan.fecha_cierre,
        encargado: plan.encargado,
        avance_plan_accion: plan.avance_plan_accion,
        efectividad: plan.efectividad,
        reuniones: plan.reuniones.map(reu => ({
          id_reunion: reu.id_reunion,
          id_plan_accion: reu.id_plan_accion,
          titulo: reu.titulo,
          proposito: reu.proposito,
          conclusiones: reu.conclusiones,
          fecha_reunion: reu.fecha_reunion,
          asistentes: reu.asistentes.map(as => ({
            id_asistente: as.id_asistente,
            nombre: as.nombre,
            cargo: as.cargo,
            email: as.email,
          })),
        })),
      })),
    }));

    res.json(data);
  } catch (error) {
    console.error("❌ Error al obtener los planes de acción:", error);
    res.status(500).json({ message: "Error al obtener los planes de acción" });
  }
  console.log("Datos recibidos:", req.body);
});
// ======================
// CREAR PLAN DE ACCIÓN + REUNIONES + ASISTENTES
// ======================
app.post("/plan-accion", async (req, res) => {
  console.log("📥 Solicitud recibida:", req.body);

  const { reuniones, tipo_ticket, ...planData } = req.body; // Extrae reuniones y tipo_ticket
  const t = await sequelize.transaction();

  try {
    // Importa el modelo dinámico
    const getActionPlanModel = require("./models/DynamicActionPlan");
    const PlanAccion = await getActionPlanModel();

    // 🧠 Crear plan con todos los campos que existan en el body
    const nuevoPlan = await PlanAccion.create({
      ...planData,
      tipo_ticket, // Asegúrate de guardar el tipo_ticket
    }, { transaction: t });

    // Si vienen reuniones, crearlas igual que antes
    if (Array.isArray(reuniones) && reuniones.length > 0) {
      const { Reunion, Asistente } = require("./models"); // ajusta según tu estructura real
      for (const reunionData of reuniones) {
        const { titulo, proposito, conclusiones, fecha_reunion, asistentes } = reunionData;

        const nuevaReunion = await Reunion.create(
          {
            id_plan_accion: nuevoPlan.id_plan_accion,
            titulo,
            proposito,
            conclusiones,
            fecha_reunion,
          },
          { transaction: t }
        );

        if (Array.isArray(asistentes) && asistentes.length > 0) {
          for (const asistenteData of asistentes) {
            const [asistente] = await Asistente.findOrCreate({
              where: { email: asistenteData.email },
              defaults: asistenteData,
              transaction: t,
            });
            await nuevaReunion.addAsistente(asistente, { transaction: t });
          }
        }
      }
    }

    await t.commit();
    res.status(201).json({
      message: "✅ Plan de acción creado con éxito",
      plan: nuevoPlan,
    });
  } catch (error) {
    await t.rollback();
    console.error("❌ Error al crear plan de acción:", error);
    res.status(500).json({
      message: "Error al crear plan de acción",
      error: error.message,
    });
  }
});


// ======================
// EDITAR PLAN DE ACCIÓN (Flexible)
// ======================
app.put("/plan-accion/:id", async (req, res) => {
  const { id } = req.params;
  const { reuniones, tipo_ticket, ...planData } = req.body; // Extrae reuniones y tipo_ticket
  const t = await sequelize.transaction();

  try {
    const plan = await PlanAccion.findByPk(id, { transaction: t });
    if (!plan) {
      await t.rollback();
      return res.status(404).json({ message: "Plan de acción no encontrado" });
    }

    // Actualiza TODOS los campos que hayan llegado en planData
    await plan.update({
      ...planData,
      tipo_ticket, // Asegúrate de actualizar el tipo_ticket
    }, { transaction: t });

    // Procesa reuniones (crear o actualizar según si tiene id_reunion o no)
    if (Array.isArray(reuniones) && reuniones.length > 0) {
      const { Reunion, Asistente } = require("./models");
      for (const reunionData of reuniones) {
        const { id_reunion, titulo, proposito, conclusiones, fecha_reunion, asistentes } = reunionData;
        let reunion;

        if (id_reunion) {
          reunion = await Reunion.findByPk(id_reunion, { transaction: t });
          if (!reunion) {
            throw new Error(`Reunión con id ${id_reunion} no encontrada`);
          }
          await reunion.update({ titulo, proposito, conclusiones, fecha_reunion }, { transaction: t });
        } else {
          reunion = await Reunion.create(
            {
              id_plan_accion: plan.id_plan_accion,
              titulo,
              proposito,
              conclusiones,
              fecha_reunion,
            },
            { transaction: t }
          );
        }

        // Procesa asistentes si hay
        if (Array.isArray(asistentes) && asistentes.length > 0) {
          for (const asistenteData of asistentes) {
            const [asistente] = await Asistente.findOrCreate({
              where: { email: asistenteData.email },
              defaults: asistenteData,
              transaction: t,
            });
            await reunion.addAsistente(asistente, { transaction: t });
          }
        }
      }
    }

    await t.commit();
    res.json({ message: "✅ Plan de acción actualizado correctamente" });
  } catch (error) {
    await t.rollback();
    console.error("❌ Error al actualizar plan de acción:", error);
    res.status(500).json({
      message: "Error al actualizar plan de acción",
      error: error.message,
    });
  }
});

app.post("/plan-accion/:planId/reunion", async (req, res) => {
  const { planId } = req.params;
  const reunionData = req.body;  // Incluye asistentes
  const t = await sequelize.transaction();

  console.log("📥 Solicitud a /plan-accion/:planId/reunion:", req.body);  // Log de entrada

  try {
    const plan = await PlanAccion.findByPk(planId, { transaction: t });
    if (!plan) {
      await t.rollback();
      console.error(`Plan con ID ${planId} no encontrado`);
      return res.status(404).json({ message: "Plan no encontrado" });
    }

    const nuevaReunion = await Reunion.create(
      {
        id_plan_accion: planId,
        titulo: reunionData.titulo,
        proposito: reunionData.proposito,
        conclusiones: reunionData.conclusiones,
        fecha_reunion: reunionData.fecha_reunion,
      },
      { transaction: t }
    );

    if (Array.isArray(reunionData.asistentes) && reunionData.asistentes.length > 0) {
      for (const asistenteData of reunionData.asistentes) {
        console.log(`Agregando asistente:`, asistenteData);  // Log por asistente
        const [asistente] = await Asistente.findOrCreate({
          where: { email: asistenteData.email },
          defaults: asistenteData,
          transaction: t,
        });
        await nuevaReunion.addAsistente(asistente, { transaction: t });
      }
    }

    await t.commit();
    console.log(`✅ Reunión creada para plan ID ${planId}`);
    res.status(201).json(nuevaReunion);
  } catch (error) {
    await t.rollback();
    console.error("❌ Error detallado al crear reunión:", error);  // Log detallado
    res.status(500).json({ message: "Error al crear reunión", error: error.message });
  }
});

app.put("/plan-accion/:planId/reunion/:reunionId", async (req, res) => {
  const planId = parseInt(req.params.planId);  // Convierte a número
  const reunionId = parseInt(req.params.reunionId);  // Convierte a número
  const reunionData = req.body;
  const t = await sequelize.transaction();  // Inicia transacción

  console.log(`[DEBUG] Solicitud PUT recibida para planId: ${planId}, reunionId: ${reunionId}, Body:`, req.body);

  try {
    const reunion = await Reunion.findByPk(reunionId, { transaction: t });  // Busca la reunión
    console.log(`[DEBUG] Reunión encontrada en DB:`, reunion);

    if (!reunion) {
      console.log(`[DEBUG] Reunión con ID ${reunionId} no encontrada en la DB`);
      await t.rollback();
      return res.status(404).json({ message: "Reunión no encontrada" });
    }

    if (reunion.id_plan_accion !== planId) {  // Comparación corregida
      console.log(`[DEBUG] ID de plan no coincide: Esperado ${planId}, Encontrado ${reunion.id_plan_accion}`);
      await t.rollback();
      return res.status(404).json({ message: "Reunión no pertenece a este plan" });
    }

    // Actualiza los campos de la reunión
    await reunion.update(reunionData, { transaction: t });

    if (Array.isArray(reunionData.asistentes)) {
      const actualesAsistentes = await reunion.getAsistentes({ transaction: t });
      const nuevosAsistentesEmails = reunionData.asistentes.map(a => a.email);

      // Elimina asistentes no presentes
      for (const asistente of actualesAsistentes) {
        if (!nuevosAsistentesEmails.includes(asistente.email)) {
          await reunion.removeAsistente(asistente, { transaction: t });
        }
      }

      // Agrega o actualiza asistentes
      for (const asistenteData of reunionData.asistentes) {
        const [asistente] = await Asistente.findOrCreate({
          where: { email: asistenteData.email },
          defaults: asistenteData,
          transaction: t,
        });
        await reunion.addAsistente(asistente, { transaction: t });
      }
    }

    await t.commit();
    res.json(reunion);  // Devuelve la reunión actualizada
  } catch (error) {
    await t.rollback();
    console.error("❌ Error al actualizar reunión:", error);
    res.status(500).json({ message: "Error al actualizar reunión", error: error.message });
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
