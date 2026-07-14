const express = require("express");
const router = express.Router();

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const multer = require("multer");


// ----------------------------
// MULTER STORAGE (NEWS IMAGES)
// ----------------------------
const newsStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, "../newsEvents/news"));
    },
    filename: function (req, file, cb) {
        const uniqueName = Date.now() + "_" + file.originalname.replace(/\s+/g, "_");
        cb(null, uniqueName);
    }
});

const uploadNews = multer({ storage: newsStorage });


// ----------------------------
// GET NEWS
// ----------------------------
router.get("/get-news", (req, res) => {

    const filePath = path.join(__dirname, "../news.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ rows: [] });
    }

    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    res.json({ rows });
});


// ----------------------------
// ADD NEWS (with image upload)
// ----------------------------
router.post("/add-news", uploadNews.single("image"), (req, res) => {

    const { title, date, description } = req.body;

    const filePath = path.join(__dirname, "../news.xlsx");

    if (!title || !date || !description) {
        return res.json({ success: false, message: "Missing fields" });
    }

    // Image path to store in Excel
    let imagePath = "";
    if (req.file) {
        imagePath = "newsEvents/news/" + req.file.filename;
    }

    // Load existing rows or create header
    let rows = [];
    if (fs.existsSync(filePath)) {
        const wb = XLSX.readFile(filePath);
        const sheet = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    } else {
        rows = [["ID", "Title", "Date", "Description", "Image"]];
    }

    // Generate new ID
    let maxNum = 100;
    for (let i = 1; i < rows.length; i++) {
        let rid = String(rows[i][0] || "");
        if (rid.startsWith("N")) {
            let num = parseInt(rid.substring(1));
            if (!isNaN(num) && num > maxNum) maxNum = num;
        }
    }
    const newId = "N" + (maxNum + 1);

    rows.push([newId, title, date, description, imagePath]);

    const wbNew = XLSX.utils.book_new();
    const wsNew = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wbNew, wsNew, "Sheet1");
    XLSX.writeFile(wbNew, filePath);

    res.json({ success: true, message: "News added successfully", id: newId });
});


// ----------------------------
// DELETE NEWS
// ----------------------------
router.post("/delete-news", (req, res) => {

    const { id } = req.body;
    const filePath = path.join(__dirname, "../news.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ success: false, message: "news.xlsx not found" });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    let newRows = rows.filter((r, index) => {
        if (index === 0) return true;
        return String(r[0] || "").trim() !== String(id || "").trim();
    });

    const newSheet = XLSX.utils.aoa_to_sheet(newRows);   // ✅ MUST be newRows
    wb.Sheets[wb.SheetNames[0]] = newSheet;
    XLSX.writeFile(wb, filePath);

    res.json({ success: true, message: "News deleted successfully" });
});

module.exports = router;








;