const jwt = require("jsonwebtoken");

function generateToken(payload) {
  return jwt.sign(payload, process.env.JWT_SECRET || "secret-key", {
    expiresIn: "8h",
  });
}

module.exports = { generateToken };
