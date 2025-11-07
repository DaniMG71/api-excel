const jwt = require("jsonwebtoken");
const secretKey = process.env.JWT_SECRET || "secret-key";

const authorize = (rolesPermitidos = []) => {
  return (req, res, next) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader) {
        return res.status(401).json({ error: "Token no proporcionado" });
      }

      const token = authHeader.split(" ")[1];
      const decoded = jwt.verify(token, secretKey);

      // Guardar info del usuario en la request
      req.user = decoded;

      // Verificar rol si aplica
      if (rolesPermitidos.length > 0 && !rolesPermitidos.includes(decoded.role)) {
        return res.status(403).json({
          error: `Acceso denegado: el rol '${decoded.role}' no tiene permiso para esta acción.`,
        });
      }

      next();
    } catch (error) {
      console.error("❌ Error en autorización:", error.message);

      if (error.name === "TokenExpiredError") {
        return res.status(401).json({ error: "La sesión ha expirado. Por favor, inicia sesión nuevamente." });
      }

      if (error.name === "JsonWebTokenError") {
        return res.status(401).json({ error: "Token inválido. Acceso no autorizado." });
      }

      return res.status(500).json({ error: "Error al procesar la autorización." });
    }
  };
};

module.exports = authorize;
