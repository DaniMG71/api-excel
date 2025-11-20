const { authenticate } = require("ldap-authentication");
const getUser = require("../models/DynamicUsers");
const { generateToken } = require("../services/jwt.services");
 
// ======================
// CONFIGURACIÓN LDAP
// ======================
const LDAP_URL = process.env.LDAP_URL;
const LDAP_BASE_DN = process.env.LDAP_BASE_DN;
const LDAP_DOMAIN = process.env.LDAP_DOMAIN;
 
const authLogin = async (req, res) => {
  const { username, password } = req.body;
 
  try {
    // 🧩 Validar entrada
    if (!username || !password) {
      return res.status(400).json({ error: "Usuario y contraseña requeridos" });
    }
 
    // 🔐 Autenticación LDAP
    const ldapUser = await authenticate({
      ldapOpts: { url: LDAP_URL },
      userDn: `${username}@${LDAP_DOMAIN}`,
      userPassword: password,
      userSearchBase: LDAP_BASE_DN,
      usernameAttribute: "sAMAccountName",
      username,
      attributes: ["sAMAccountName", "displayName", "memberOf", "mail"],
    });
 
    if (!ldapUser) {
      return res.status(404).json({ error: "Usuario no válido" });
    }
 
    // ⚙️ Modelos
    const [User] = await Promise.all([getUser()]);
 
    // 🔎 Buscar usuario en BD
    let user = await User.findOne({ where: { username } });
 
    if (!user) {
      return res.status(401).json({
        error: "Usuario no registrado, contactese con el administrador",
      });
    }
 
    // 🆕 COMPLETAR CAMPOS VACÍOS: full_name y email
    let updated = false;
 
    const ldapFullName = ldapUser.displayName || null;
    const ldapEmail = ldapUser.mail || `${username}@${LDAP_DOMAIN}`;
 
    if (!user.nombre || user.nombre.trim() === "") {
      user.nombre = ldapFullName;
      updated = true;
    }
 
    if (!user.email || user.email.trim() === "") {
      user.email = ldapEmail;
      updated = true;
    }

    user.last_login = new Date();  // Guarda la fecha/hora actual
    updated = true;
 
    if (updated) {
      await user.save(); // 💾 Actualiza en la base de datos
      console.log(`🔄 Usuario actualizado con datos LDAP: ${username}`);
    }
 
    // 🎫 Generar token
    const token = generateToken({
      user_id: user.user_id,
      username: user.username,
      role: user.role,
    });
 
    console.log(`✅ Token generado para ${user.role}: ${username}`);
    console.log("Datos de LDAP obtenidos para usuario:", {
  username,
  nombre: ldapFullName,
  email: ldapUser.mail
});
 
    // 📤 Respuesta final
    return res.status(200).json({
      message: `Autenticación exitosa (${user.role})`,
      user: {
        sAMAccountName: ldapUser.sAMAccountName || user.username,
        displayName: ldapUser.displayName,
        dn: ldapUser.dn || `${ldapUser.sAMAccountName}@midominio.com`,
        role: user.role,
        user_id: user.user_id,
        full_name: user.full_name,
        email: user.email,
      },
      token,
      
    });
    
  } catch (err) {
    return res.status(401).json({
      error: "Credenciales inválidas o error de autenticación",
      details: err.message,
    });
  }
};
 
module.exports = { authLogin };
 
 