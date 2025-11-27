const express = require("express");
const router = express.Router();
const { authLogin } = require("../controllers/auth.controller");

router.post("/login", authLogin);

module.exports = router;