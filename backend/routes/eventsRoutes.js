const express = require("express");
const router = express.Router();

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const multer = require("multer");


// ----------------------------
// MULTER STORAGE (EVENT IMAGES)
// ----------------------------
const eventStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, path.join(__dirname, "../newsEvents/events"));
    },
    filename: function (req, file, cb) {
        const uniqueName = Date.now() + "_" + file.originalname.replace(/\s+/g, "_");
        cb(null, uniqueName);
    }
});

const uploadEvent = multer({ storage: eventStorage });


// ----------------------------
// GET EVENTS
// ----------------------------
router.get("/get-events", (req, res) => {

    const filePath = path.join(__dirname, "../events.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ rows: [] });
    }

    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    res.json({ rows });
});


// ----------------------------
// ADD EVENT (with image upload)
// ----------------------------
router.post("/add-event", uploadEvent.single("image"), (req, res) => {

    const { title, date, description } = req.body;
    const filePath = path.join(__dirname, "../events.xlsx");

    if (!title || !date || !description) {
        return res.json({ success: false, message: "Missing fields" });
    }

    let imagePath = "";
    if (req.file) {
        imagePath = "newsEvents/events/" + req.file.filename;
    }

    let rows = [];
    if (fs.existsSync(filePath)) {
        const wb = XLSX.readFile(filePath);
        const sheet = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    } else {
        rows = [["ID", "Title", "Date", "Description", "Image"]];
    }

    let maxNum = 100;
    for (let i = 1; i < rows.length; i++) {
        let rid = String(rows[i][0] || "");
        if (rid.startsWith("E")) {
            let num = parseInt(rid.substring(1));
            if (!isNaN(num) && num > maxNum) maxNum = num;
        }
    }

    const newId = "E" + (maxNum + 1);

    rows.push([newId, title, date, description, imagePath]);

    const wbNew = XLSX.utils.book_new();
    const wsNew = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wbNew, wsNew, "Sheet1");
    XLSX.writeFile(wbNew, filePath);

    res.json({ success: true, message: "Event added successfully", id: newId });
});


// ----------------------------
// DELETE EVENT
// ----------------------------

router.post("/update-event", (req, res) => {

    const { id, title, date, description } = req.body;
    const filePath = path.join(__dirname, "../events.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ success: false, message: "events.xlsx not found" });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    let found = false;

    for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(id).trim()) {
            rows[i][1] = title;
            rows[i][2] = date;
            rows[i][3] = description;
            found = true;
            break;
        }
    }

    if (!found) {
        return res.json({ success: false, message: "Event ID not found" });
    }

    const newSheet = XLSX.utils.aoa_to_sheet(rows);
    wb.Sheets[wb.SheetNames[0]] = newSheet;
    XLSX.writeFile(wb, filePath);

    res.json({ success: true, message: "Event updated" });
});


// ----------------------------
// DELETE EVENT
// ----------------------------
router.post("/delete-event", (req, res) => {

    const { id } = req.body;
    const filePath = path.join(__dirname, "../events.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ success: false, message: "events.xlsx not found" });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    let newRows = rows.filter((r, index) => {
        if (index === 0) return true;
        return String(r[0] || "").trim() !== String(id || "").trim();
    });

    const newSheet = XLSX.utils.aoa_to_sheet(newRows);
    wb.Sheets[wb.SheetNames[0]] = newSheet;
    XLSX.writeFile(wb, filePath);

    res.json({ success: true, message: "Event deleted successfully" });
});


module.exports = router;