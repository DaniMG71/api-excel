const express = require("express");
const router = express.Router();
const { syncUser } = require("../controllers/user.controller.js");

router.post("/sync", syncUser);

module.exports = router;
