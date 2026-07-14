const express = require("express");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const router = express.Router();

// ---------------------------
// GET WELCOME
// ---------------------------
router.get("/get-welcome", (req, res) => {

    const filePath = path.join(__dirname, "../get-welcome.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ rows: [] });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    res.json({ rows });
});

// ---------------------------
// DELETE WELCOME
// ---------------------------
router.post("/delete-welcome", (req, res) => {

    const { rowIndex } = req.body;

    const filePath = path.join(__dirname, "../get-welcome.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ success: false });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    // rowIndex must not delete header
    if (rowIndex > 0 && rowIndex < rows.length) {
        rows.splice(rowIndex, 1);
    }

    const newWb = XLSX.utils.book_new();
    const newWs = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(newWb, newWs, "Welcome");
    XLSX.writeFile(newWb, filePath);

    res.json({ success: true });
});

module.exports = router;