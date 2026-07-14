const express = require("express");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const router = express.Router();

// ---------------------------
// Paths
// ---------------------------
const globalFilePath = path.join(__dirname, "../global_members.xlsx");
const memberFolderPath = path.join(__dirname, "../memberfiles");

// Ensure memberfiles folder exists
if (!fs.existsSync(memberFolderPath)) {
    fs.mkdirSync(memberFolderPath);
}

// Ensure global file exists
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

// Date format mm/dd/yyyy
function getCurrentDate() {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
}


// ---------------------------
// LOAD PERSONAL FILE
// ---------------------------
router.post("/load-personal", (req, res) => {

    const { username } = req.body;

    let members = readGlobal();
    const user = members.find(u => u.username === username);

    if (!user) return res.json({ rows: [] });

    const personalPath = path.join(
        memberFolderPath,
        `${user.firstName}_${user.lastName}.xlsx`
    );

    if (!fs.existsSync(personalPath))
        return res.json({ rows: [] });

    const wb = XLSX.readFile(personalPath);
    const ws = wb.Sheets["Personal"];

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    res.json({ rows });
});


// ---------------------------
// SAVE PERSONAL FILE
// ---------------------------
router.post("/save-personal", (req, res) => {

    try {
        const { username, rows } = req.body;

        if (!rows || rows.length === 0) {
            return res.json({ message: "No data received" });
        }

        let members = readGlobal();
        const user = members.find(u => u.username === username);

        if (!user) return res.json({ message: "User not found" });

        const personalPath = path.join(
            memberFolderPath,
            `${user.firstName}_${user.lastName}.xlsx`
        );

        const newWs = XLSX.utils.aoa_to_sheet(rows);

        // formatting
        newWs["!cols"] = [
            { wch: 15 },   // Date
            { wch: 120 }   // Activity
        ];

        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, newWs, "Personal");

        XLSX.writeFile(newWb, personalPath);

        res.json({ message: "Personal file saved" });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Save error" });
    }
});


// ---------------------------
// APPEND ACTIVITY
// ---------------------------
router.post("/append", (req, res) => {

    try {
        const { username, data } = req.body;

        let members = readGlobal();
        const user = members.find(u => u.username === username);

        if (!user)
            return res.json({ message: "User not found." });

        const personalPath = path.join(
            memberFolderPath,
            `${user.firstName}_${user.lastName}.xlsx`
        );

        let activityRows = [];

        if (fs.existsSync(personalPath)) {

            const wb = XLSX.readFile(personalPath);
            const ws = wb.Sheets["Personal"];

            const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

            if (rows.length >= 4) {
                activityRows = rows.slice(3);
            }
        }

        // Add new row
        activityRows.push([
            getCurrentDate(),
            data
        ]);

        const newSheetData = [
            [user.firstName, user.lastName, user.username, user.EmailAddress, user.memberCode],
            [],
            ["Date", "Activity"],
            ...activityRows
        ];

        const newWs = XLSX.utils.aoa_to_sheet(newSheetData);

        newWs["!cols"] = [
            { wch: 15 },
            { wch: 80 }
        ];

        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, newWs, "Personal");

        XLSX.writeFile(newWb, personalPath);

        res.json({ message: "Data appended successfully." });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Append error." });
    }
});

module.exports = router;