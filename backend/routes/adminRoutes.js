const express = require("express");
const router = express.Router();

// ---------------------------
// ADMIN LOGIN FIXED USER/PASS
// ---------------------------
const ADMIN_USERNAME = "admin";
const ADMIN_PASSWORD = "admin123";   // change later

// ---------------------------
// ADMIN LOGIN ROUTE
// ---------------------------
router.post("/admin-login", (req, res) => {

    const { username, password } = req.body;

    if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
        return res.json({ success: true });
    }

    res.json({ success: false, message: "Invalid Admin Login" });
});

module.exports = router;