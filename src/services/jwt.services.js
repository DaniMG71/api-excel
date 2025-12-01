const jwt = require("jsonwebtoken");

function generateToken(payload) {
  let expiresIn;

  switch (payload.role) {
    case "admin":
      expiresIn = "4h"; // Admin: 4 horas
      break;
    case "user":
      expiresIn = "30m"; // User: 30 minutos
      break;
    case "superadmin":
      expiresIn = "5h"; 
    default:
      expiresIn = "5m"; // Valor por defecto
      break;
  }

  return jwt.sign(payload, process.env.JWT_SECRET || "secret-key", {
    expiresIn,
  });
}

module.exports = { generateToken };
