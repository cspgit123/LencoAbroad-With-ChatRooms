const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const router = express.Router();

// ---------------------------
// Upload Folder Setup
// ---------------------------
const uploadsFolder = path.join(__dirname, "../../uploads");

// create uploads folder if missing
if (!fs.existsSync(uploadsFolder)) {
    fs.mkdirSync(uploadsFolder);
}

// ---------------------------
// Folder size calculator
// ---------------------------
function getFolderSize(folderPath) {
    let total = 0;

    if (!fs.existsSync(folderPath)) return 0;

    const files = fs.readdirSync(folderPath);

    for (const file of files) {
        const fullPath = path.join(folderPath, file);
        const stat = fs.statSync(fullPath);

        if (stat.isDirectory()) {
            total += getFolderSize(fullPath);
        } else {
            total += stat.size;
        }
    }

    return total;
}

// ---------------------------
// Read global members for storage limit
// ---------------------------
const globalFilePath = path.join(__dirname, "../global_members.xlsx");

function ensureGlobalFile() {
    if (!fs.existsSync(globalFilePath)) {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(wb, ws, "Members");
        XLSX.writeFile(wb, globalFilePath);
    }
}

function readGlobal() {
    ensureGlobalFile();
    const wb = XLSX.readFile(globalFilePath);
    const ws = wb.Sheets["Members"];
    return XLSX.utils.sheet_to_json(ws);
}

function getUserStorageLimit(username) {

    let members = readGlobal();
    const user = members.find(u => u.username === username);

    // default limit
    if (!user) return 200 * 1024 * 1024; // 200MB

    const mb = parseInt(user.storageLimitMB);

    if (!mb || mb <= 0) {
        return 200 * 1024 * 1024;
    }

    return mb * 1024 * 1024;
}

// ---------------------------
// Multer Setup
// ---------------------------
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsFolder);
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + "_" + file.originalname);
    }
});

const upload = multer({
    storage: storage,
    limits: { fileSize: 20 * 1024 * 1024 } // 20MB limit
});

// ---------------------------
// UPLOAD ROUTE
// ---------------------------
router.post("/upload", upload.single("media"), (req, res) => {

    try {
        const username = req.body.username;

        if (!username) {
            return res.json({ success: false, message: "Username missing in upload request" });
        }

        if (!req.file) {
            return res.json({ success: false, message: "No file uploaded" });
        }

        const usedBytes = getFolderSize(uploadsFolder);
        const maxBytes = getUserStorageLimit(username);
        const newFileSize = req.file.size;

        if (usedBytes + newFileSize > maxBytes) {

            // delete file immediately
            fs.unlinkSync(req.file.path);

            return res.json({
                success: false,
                message: "Storage limit exceeded. Please delete old uploads or contact admin."
            });
        }

        res.json({
            success: true,
            message: "Upload successful",
            file: req.file.filename
        });

    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, message: "Upload error" });
    }
});

module.exports = router;