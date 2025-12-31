# api-excel — Documentación técnica centrada en index.js

Índice
- Visión general
- Tecnologías principales
- Requisitos previos
- Variables de entorno
- Instalación y ejecución
- Flujo de una solicitud HTTP (desde index.js)
- Cómo se importan/definen y utilizan los modelos Sequelize
- Operaciones ORM principales (ejemplos)
- Middlewares relevantes
- Endpoints (detallados)
  - /etl (definida en index.js)
  - Endpoints desde auth.routes.js y user.routes.js
- Ejemplos prácticos (curl)
- Notas de seguridad y troubleshooting

-------------------------
Visión general
-------------------------
Este proyecto implementa una API REST en Node.js cuya finalidad principal es la ingesta y procesamiento de datos (por ejemplo, alimentación desde archivos CSV/Excel) y la exposición de endpoints autenticados mediante LDAP y JWT. El fichero central `index.js` configura la aplicación Express, registra middlewares (CORS, JSON, manejo de archivos), monta rutas y contiene lógica de ETL (extracción y persistencia de datos) con uso de transacciones Sequelize para garantizar atomicidad.

Propósito: recibir solicitudes HTTP (incluyendo cargas de archivos), validar autorización, transformar/parsear datos y persistirlos en una base de datos relacional usando Sequelize. Además, provee endpoints para autenticación LDAP y sincronización de usuario.

-------------------------
Tecnologías principales
-------------------------
- Node.js (runtime)
- Express (servidor HTTP)
- Sequelize (ORM)
- PostgreSQL / MSSQL / otros (a través del dialecto Sequelize)
- jsonwebtoken (JWT)
- ldap-authentication / ldapjs (autenticación LDAP)
- multer (manejo de multipart/form-data en memoria)
- dotenv (gestión de variables de entorno)
- xlsx / csv parsing (librerías incluidas en package.json; el código del ETL procesa CSV en memoria)

-------------------------
Requisitos previos
-------------------------
- Node.js >= 16 (se recomienda la LTS vigente)
- Base de datos compatible (Postgres, MSSQL, etc.) accesible con credenciales
- Servidor LDAP accesible si se utilizará la autenticación LDAP
- Variables de entorno correctamente definidas (sección siguiente)

-------------------------
Variables de entorno
-------------------------
El comportamiento del proyecto depende de variables de entorno. Las principales son:

- DB_NAME — nombre de la base de datos
- DB_USER — usuario de BD
- DB_PASS — contraseña de BD
- DB_HOST — host de BD
- DB_PORT — puerto de BD
- DB_DIALECT — dialecto para Sequelize (ej. "postgres", "mssql")
- LDAP_URL — URL del servidor LDAP
- LDAP_BASE_DN — base DN para búsquedas LDAP
- LDAP_DOMAIN — dominio para bind (ej. midominio.com)
- JWT_SECRET — clave secreta para firmar JWT
- PORT — puerto donde escuchará el servidor (por defecto, puede ser 3000 si no está definido)

Coloque estas variables en un archivo `.env` en la raíz del proyecto antes de iniciar la aplicación.

-------------------------
Instalación y ejecución
-------------------------
1. Clonar el repositorio y situarse en la carpeta raíz.
2. Instalar dependencias:
   npm install
3. Crear archivo `.env` con las variables reseñadas.
4. Iniciar el servidor:
   node index.js
   (Puede usar nodemon u otros gestores durante desarrollo.)

Si necesita ejecutar en entorno productivo, configure un proceso manager (pm2, systemd) y asegure las variables de entorno.

-------------------------
Flujo básico de una solicitud HTTP (desde index.js)
-------------------------
1. Cliente envía una solicitud HTTP al servidor (por ejemplo POST /etl o POST /auth/login).
2. Entrada en Express:
   - El objeto `app` creado en `index.js` pasa la solicitud por los middlewares registrados en orden (CORS, parsing de multipart con multer para endpoints que lo requieren, express.json para cuerpos JSON, middleware de autorización cuando se aplica, etc.).
3. Middleware de autorización (si la ruta lo requiere):
   - Extrae el encabezado `Authorization: Bearer <token>`.
   - Verifica el token JWT usando `JWT_SECRET`.
   - Valida que el rol del usuario esté permitido para la ruta (ej.: `authorize(['superadmin'])`).
   - Si es válido, adjunta `req.user = decoded` y permite continuar.
4. Enrutamiento:
   - Si la ruta está implementada en un router (por ejemplo `auth.routes.js`), Express delega allí.
   - Si la ruta está definida directamente en `index.js` (por ejemplo `/etl`), la función manejadora definida en `index.js` procesa la solicitud.
5. Controller / lógica:
   - La función del controller puede invocar servicios (LDAP, JWT) o interactuar con modelos Sequelize.
   - En el caso de ETL, se parsea el archivo recibido (en memoria) y se construyen registros.
6. Acceso a la base de datos con Sequelize:
   - Los modelos Sequelize (dinámicos o estáticos) ejecutan operaciones: findOne, create, bulkCreate, update, destroy, transacciones.
   - Para operaciones que afectan múltiples tablas, se usa `sequelize.transaction()` para garantizar atomicidad y permitir rollback en fallo.
7. Respuesta:
   - El controller construye una respuesta HTTP (status code y JSON) y la devuelve al cliente. Si hubo error, se devuelve un error con un código HTTP adecuado (400, 401, 403, 500, etc.).

-------------------------
Cómo se importan y usan los modelos Sequelize
-------------------------
1. Modelos dinámicos:
   - En `src/models/DynamicUsers.js`, `DynamicTicket.js`, `DynamicTienda.js`, etc., cada fichero exporta una función asíncrona (por ejemplo `getUserModel()`).
   - Esa función ejecuta una consulta a `information_schema.columns` para obtener la estructura de la tabla en la BD.
   - A partir del resultado, construye `modelDefinition` mapeando tipos SQL a `DataTypes` de Sequelize y llama a `sequelize.define('Name', modelDefinition, { tableName: '...', timestamps: false })`.
   - Ejemplo de mapeo: `character varying` → `DataTypes.STRING`, `timestamp without time zone` → `DataTypes.DATE`.
   - Ventaja: si el esquema de la BD varía o se actualiza, el modelo se reconstruye según la estructura actual.
   - Uso: en `index.js` o en controllers se invoca `const User = await getUserModel();` y luego `User.findOne(...)`, `User.create(...)`, etc.
2. Modelos estáticos:
   - `src/models/TicketTienda.js` define un modelo de forma estática con campos y opciones (índices, unique).
   - Estos modelos se importan directamente con `require('./src/models/TicketTienda')` y se usan inmediatamente.
3. Importación en runtime:
   - Dado que la construcción de modelos dinámicos es asíncrona, el código usa `await getModel()` o `getModel().then(model => ...)` antes de operar con el modelo.
   - En `index.js` hay patrones que cargan los modelos al inicio para uso posterior (por ejemplo `getTiendaModel().then(model => Tienda = model)`).

-------------------------
Qué operaciones realiza el ORM (Sequelize)
-------------------------
- Definición dinámica de modelos (sequelize.define).
- Consultas de lectura: findOne, findAll, queries directas con sequelize.query.
- Inserciones: create, bulkCreate.
- Actualizaciones: instance.save() o Model.update(...).
- Borrado: Model.destroy(...).
- Transacciones: `const t = await sequelize.transaction()` y luego pasar `{ transaction: t }` a las operaciones; `t.commit()` / `t.rollback()` ante éxito/fallo.
- Índices y constraints (en modelos estáticos) para optimizar búsquedas y evitar duplicados (ej.: índice único en TicketTienda).
- El código también ejecuta queries directas a `information_schema` para introspección del esquema.

-------------------------
Middlewares relevantes
-------------------------
- CORS (habilitado globalmente con `app.use(cors())`).
- multer: configuración `multer({ storage: multer.memoryStorage() })` para recibir archivos en memoria. Importante: multer debe registrarse antes de `express.json()` para evitar conflictos con multipart.
- autorización (`src/middlewares/authorize.js`): valida JWT, maneja errores (`TokenExpiredError`, `JsonWebTokenError`) y verifica roles permitidos.
- express.json (para JSON bodies) — usado después de configurar multer.

-------------------------
Endpoints definidos (detallados)
-------------------------
A continuación se detallan los endpoints expuestos por el conjunto de archivos presentes. index.js importa routers desde `src/routes/auth.routes.js` y `src/routes/user.routes.js`, y define directamente la ruta `/etl`. En este documento se asume que los routers se montan en `/auth` y `/user` respectivamente (si la aplicación monta los routers en otra ruta, ajustar las rutas a la configuración real en `index.js`).

1) POST /etl
- Ruta: /etl
- Método: POST
- Protección: require rol "superadmin". Middleware: authorize(['superadmin'])
- Tipo de contenido: multipart/form-data
- Parámetros:
  - file (form field): el archivo CSV que será procesado. En el servidor `multer` lo almacena en memoria y está disponible en `req.file`.
  - Authorization header: `Authorization: Bearer <token>` (token JWT con rol `superadmin`).
- Comportamiento:
  - Valida la presencia del archivo; si no existe devuelve 400.
  - Concatena/convierte `req.file.buffer` a string UTF-8 y parsea manualmente el CSV (separado por comas, primera línea headers).
  - Convierte cada línea en un objeto con claves desde los headers.
  - Inicia una transacción Sequelize (`sequelize.transaction()`).
  - Inserta/actualiza datos en una o varias tablas según la lógica interna (insertado en tablas dinámicas y relaciones con `TicketTienda` cuando aplique). Si ocurre un error se hace rollback y se devuelve error 500.
  - Retorna un JSON indicando cuantos registros se procesaron y el resultado (200 o error).
- Observaciones:
  - El parsing CSV es simple y depende de separador coma; no maneja comillas complejas ni escapes avanzados.
  - Idealmente en producción usar una librería CSV robusta si los archivos contienen comas embebidas.

2) POST /auth/login
- Ruta (router): definido en `src/routes/auth.routes.js` como `router.post("/login", authLogin);`
- Ruta completa esperada: /auth/login (si se monta en /auth)
- Método: POST
- Tipo de contenido: application/json
- Body:
  - username (string) — nombre de usuario LDAP (sAMAccountName)
  - password (string) — contraseña
- Comportamiento:
  - Valida que username y password estén presentes.
  - Autentica contra LDAP usando `ldap-authentication` (config con LDAP_URL, LDAP_BASE_DN, LDAP_DOMAIN).
  - Si la autenticación LDAP es exitosa, busca el usuario en la tabla `users` (modelo dinámico).
  - Actualiza campos faltantes (nombre, correo) si se obtuvieron de LDAP y registra `last_login`.
  - Genera un token JWT con payload { user_id, username, role } usando `src/services/jwt.services.js`. La expiración del token depende del role:
    - admin: 4 horas
    - user: 30 minutos
    - superadmin: 5 horas
    - por defecto: 5 minutos
  - Retorna: objeto con `message`, `user` (información básica) y `token`.
- Códigos relevantes:
  - 200: autenticación exitosa + token
  - 400: falta de campos
  - 401/403: usuario no registrado o credenciales inválidas
  - 500: error interno / LDAP inaccesible

3) POST /user/sync
- Ruta (router): definido en `src/routes/user.routes.js` como `router.post("/sync", syncUser);`
- Ruta completa esperada: /user/sync (si se monta en /user)
- Método: POST
- Tipo de contenido: application/json
- Body:
  - sAMAccountName (string) — requerido
  - displayName (string) — requerido
  - userDn (string) — opcional (DN LDAP)
- Comportamiento:
  - Endpoint pensado para sincronizar datos de usuario desde LDAP al sistema.
  - Valida campos mínimos (sAMAccountName y displayName).
  - Genera un token JWT de prueba para el usuario sincronizado.
  - Retorna un objeto `user` simulado y un token. Nota: en el código actual la sincronización es una simulación; para persistir en BD, implementar la lógica con `getUserModel()` y `User.create()` o `User.update()`.

4) Rutas adicionales
- Cualquier otra ruta debe revisarse en `index.js` completo. Los routers importados (`auth.routes.js` y `user.routes.js`) únicamente definen los endpoints detallados arriba.

-------------------------
Ejemplos prácticos
-------------------------
1) Login (obtener token):
curl:
  curl -X POST http://localhost:3000/auth/login \
    -H "Content-Type: application/json" \
    -d '{"username":"jdoe","password":"miPassword"}'

Respuesta esperada (ejemplo):
{
  "message": "Autenticación exitosa (user)",
  "user": { "sAMAccountName": "jdoe", "displayName": "John Doe", "role": "user", "user_id": 123, "full_name": "...", "email": "jdoe@midominio.com" },
  "token": "eyJhbGciOi..."
}

2) Subir CSV a /etl:
- Requisito: token JWT con rol `superadmin`.
curl:
  curl -X POST http://localhost:3000/etl \
    -H "Authorization: Bearer <TOKEN_SUPERADMIN>" \
    -F "file=@/ruta/al/archivo.csv"

Respuesta esperada:
- 200: { message: "Procesados N registros", details: { inserted: X, updated: Y } }
- 400: archivo no recibido o CSV vacío
- 401/403: token ausente/rol no autorizado

-------------------------
Notas de seguridad y recomendaciones
-------------------------
- JWT_SECRET debe ser una cadena fuerte y no almacenarse en repositorios.
- El endpoint /etl concede poder de inserción masiva: restringir su acceso a usuarios administrativos (ya lo protege por rol `superadmin`).
- Los datos CSV se procesan en memoria; para archivos muy grandes considerar streaming o procesamiento por lotes para evitar OOM.
- Validar y sanitizar todos los campos antes de insertarlos en la BD para prevenir inyección o datos inválidos.
- LDAP y BD deben comunicarse sobre canales seguros (LDAPS y conexiones TLS a la base de datos).

-------------------------
Resolución de problemas comunes
-------------------------
- Error de conexión DB: revisar variables DB_* y la accesibilidad desde la máquina que ejecuta la app.
- Error de autenticación LDAP: verificar LDAP_URL, LDAP_DOMAIN y que las credenciales sean correctas; revisar si la app debe usar `user@domain` o DN completo.
- Token expirado: el middleware devuelve 401 con mensaje de expiración cuando JWT expiró; solicitar nuevo login.
- `No se encontraron columnas` al construir modelos dinámicos: verificar que la tabla existe en el schema `public` y que el usuario DB tiene permisos para consultar `information_schema`.

-------------------------
Referencias de código relevantes
-------------------------
- Entrada DB: src/config/database.js
- Autenticación LDAP: src/controllers/auth.controller.js (usa ldap-authentication)
- Servicios JWT: src/services/jwt.services.js
- Middleware autorización: src/middlewares/authorize.js
- Models dinámicos: src/models/DynamicUsers.js, src/models/DynamicTicket.js, src/models/DynamicTienda.js
- Modelo estático: src/models/TicketTienda.js
- Rutas: src/routes/auth.routes.js, src/routes/user.routes.js

-------------------------
Conclusión
-------------------------
El archivo `index.js` actúa como el eje de arranque de la aplicación: configura middlewares, maneja carga de archivos, monta routers y contiene rutas críticas (como `/etl`) que integran el parseo de archivos y la persistencia en la base de datos mediante Sequelize. El diseño combina modelos dinámicos (introspección en tiempo de ejecución) con controles de autenticación basados en LDAP y sesiones con JWT, garantizando flexibilidad frente a cambios en el esquema de base de datos y control de acceso por roles.
