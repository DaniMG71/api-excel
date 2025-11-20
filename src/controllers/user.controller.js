const { generateToken } = require("../services/jwt.services.js");
 
async function syncUser(req, res) {
  try {
    const { sAMAccountName, displayName, userDn } = req.body;
 
    if (!sAMAccountName || !displayName) {
      return res.status(400).json({ error: "Faltan campos obligatorios" });
    }
 
    // Simulamos sincronización (aquí iría tu lógica con BD si existiera)
    const user = {
      user_id: sAMAccountName,
      full_name: displayName,
      dn: userDn,
    };
 
    // Generar token JWT
    const token = generateToken({
      user_id: sAMAccountName,
      name: displayName,
    });
 
    res.json({
      message: "Usuario sincronizado correctamente",
      user,
      token,
    });
  } catch (error) {
    console.error("❌ Error en syncUser:", error);
    res.status(500).json({ error: "Error interno del servidor" });
  }
}
 
module.exports = { syncUser };
 
 