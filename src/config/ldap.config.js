const dotenv = require('dotenv');
dotenv.config();

const ldapConfig = {
  url: process.env.LDAP_URL,
  baseDN: process.env.LDAP_BASE_DN,
  domain: process.env.LDAP_DOMAIN,
}

module.exports = ldapConfig;
