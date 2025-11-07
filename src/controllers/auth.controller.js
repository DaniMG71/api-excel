const getUser = require("../models/DynamicUsers");
const { authenticateLDAP } = require("../services/ldap.services");
const { generateToken } = require("../services/jwt.services");

const authLogin = async (req, res) => {
  const { username, password } = req.body;

  try {
    if (!username || !password) {
      return res.status(400).json({ error: "Usuario y contraseña requeridos" });
    }

    // 🔐 Autenticar contra LDAP
    const userLDAP = await authenticateLDAP(username, password);

    // ⚙️ Inicializar modelo dinámicamente
    const [User] = await Promise.all([getUser()]);

    // 🔎 Buscar usuario o crearlo
    let user = await User.findOne({ where: { username } });

    if (!user) {
      user = await User.create({ username, role: "user" });
      console.log(`🆕 Usuario ${username} creado con rol 'user'`);
    }

const token = generateToken({
      user_id: user.id,
      username: user.username,
      role: user.role,
    });
    console.log(`✅ Token generado para ${username}`);

      return res.status(200).json({
        message: "Autenticación exitosa",
        user: {
          user_id: user.id,
          username: user.username,
          role: user.role,
        },
        token,
      });

  } catch (error) {
    console.error("❌ Error de autenticación LDAP:", error.message);
    res.status(401).json({
      error: "Error de autenticación",
      details: error.message,
    });
  }
};

module.exports = { authLogin };
