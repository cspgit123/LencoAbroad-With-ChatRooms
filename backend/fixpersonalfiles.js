const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const globalPath = path.join(__dirname, "global_members.xlsx");
const memberFolder = path.join(__dirname, "memberfiles");

// Ensure memberfiles folder exists
if (!fs.existsSync(memberFolder)) {
    fs.mkdirSync(memberFolder);
}

if (!fs.existsSync(globalPath)) {
    console.log("global_members.xlsx not found!");
    process.exit();
}

// Read global members
const wb = XLSX.readFile(globalPath);
const ws = wb.Sheets[wb.SheetNames[0]];
let members = XLSX.utils.sheet_to_json(ws);

// Collect all memberCodes
const allCodes = members
    .map(m => (m.memberCode || "").trim())
    .filter(c => c !== "");

// Function: random recommender code (not self)
function getRandomRecommender(selfCode) {

    if (allCodes.length <= 1) return "";

    let code = "";
    do {
        code = allCodes[Math.floor(Math.random() * allCodes.length)];
    } while (code === selfCode);

    return code;
}

// Fix each member
members.forEach(m => {

    if (!m.firstName || !m.lastName) return;

    const personalFile = path.join(memberFolder, `${m.firstName}_${m.lastName}.xlsx`);

    // ----------------------------
    // Add recommenderCode if missing
    // ----------------------------
    if (!m.recommenderCode || String(m.recommenderCode).trim() === "") {
        m.recommenderCode = getRandomRecommender(m.memberCode);
    }

    // ----------------------------
    // If personal file missing → create it
    // ----------------------------
    if (!fs.existsSync(personalFile)) {

        const personalRow = [
            m.firstName || "",
            m.lastName || "",
            m.username || "",
            m.mobileNumber || "",
            m.EmailAddress || m.email || "",
            m.city || "",
            m.state || "",
            m.country || "",
            m.NativeLocation || "",
            m.degree || "",
            m.branch || "",
            m.graduationYear || "",
            m.mastersBranch || "",
            m.profession || "",
            m.EmployerBusinessName || "",
            m.spouseName || "",
            m.dob || "",
            m.memberCode || "",
            m.recommenderCode || ""
        ];

        const sheetData = [
            personalRow,
            [""],
            ["Date", "Activity"]
        ];

        const newWs = XLSX.utils.aoa_to_sheet(sheetData);
        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, newWs, "Personal");

        XLSX.writeFile(newWb, personalFile);

        console.log("Created:", `${m.firstName}_${m.lastName}.xlsx`);
        return;
    }

    // ----------------------------
    // If personal file exists → preserve activity and rebuild
    // ----------------------------
    const pwb = XLSX.readFile(personalFile);
    const pws = pwb.Sheets[pwb.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(pws, { header: 1 });

    if (!rows || rows.length === 0) {
        rows = [];
    }

    // Preserve activity rows (after row 3)
    let activityRows = [];
    if (rows.length > 3) {
        activityRows = rows.slice(3);
    }

    const personalRow = [
        m.firstName || "",
        m.lastName || "",
        m.username || "",
        m.mobileNumber || "",
        m.EmailAddress || m.email || "",
        m.city || "",
        m.state || "",
        m.country || "",
        m.NativeLocation || "",
        m.degree || "",
        m.branch || "",
        m.graduationYear || "",
        m.mastersBranch || "",
        m.profession || "",
        m.EmployerBusinessName || "",
        m.spouseName || "",
        m.dob || "",
        m.memberCode || "",
        m.recommenderCode || ""
    ];

    const newSheetData = [
        personalRow,
        [""],
        ["Date", "Activity"],
        ...activityRows
    ];

    const newWs = XLSX.utils.aoa_to_sheet(newSheetData);

    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, newWs, "Personal");

    XLSX.writeFile(newWb, personalFile);

    console.log("Fixed:", `${m.firstName}_${m.lastName}.xlsx`);
});

// ----------------------------
// Write updated global_members.xlsx
// ----------------------------
const newGlobalWs = XLSX.utils.json_to_sheet(members);
const newGlobalWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newGlobalWb, newGlobalWs, "Members");
XLSX.writeFile(newGlobalWb, globalPath);

console.log("✅ All personal files fixed/created!");
console.log("✅ global_members.xlsx updated with recommenderCode!");