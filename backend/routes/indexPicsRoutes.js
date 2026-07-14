const express = require("express");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const router = express.Router();

// ---------------------------
// GET INDEX PICS
// ---------------------------
router.get("/get-indexpics", (req, res) => {

    const filePath = path.join(__dirname, "../indexPics.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ rows: [] });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    res.json({ rows });
});

module.exports = router;