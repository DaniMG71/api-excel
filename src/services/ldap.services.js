const ldap = require("ldapjs");

const authenticate = (username, password) => {
  return new Promise((resolve, reject) => {
    const client = ldap.createClient({
      url: process.env.LDAP_URL,
    });

    // Intenta con formato usuario@dominio
    const userPrincipalName = `${username}@${process.env.LDAP_DOMAIN}`;

    console.log(`Intentando autenticar: ${userPrincipalName}`);

    client.bind(userPrincipalName, password, (err) => {
      if (err) {
        console.error("❌ Error en autenticación LDAP:", err.message);
        client.unbind();
        return reject(new Error("Credenciales inválidas o error en LDAP."));
      }

      console.log("✅ Autenticación LDAP exitosa");
      client.unbind();
      resolve({ username });
    });
  });
};

module.exports = { authenticate };
