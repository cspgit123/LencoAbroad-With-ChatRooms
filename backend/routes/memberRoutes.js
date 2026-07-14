const express = require("express");
const XLSX = require("xlsx");
const bcrypt = require("bcrypt");
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



function cleanText(val) {
    return (val || "")
        .toString()
        .replace(/\u00A0/g, " ")          // non-breaking space
        .replace(/[\u200B-\u200D\uFEFF]/g, "") // zero-width chars
        .replace(/[^\x20-\x7E]+/g, "")    // remove non-standard symbols
        .trim();
}




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

//function writeGlobal(data) {
 //   const wb = XLSX.utils.book_new();
  //  const ws = XLSX.utils.json_to_sheet(data);
  //  XLSX.utils.book_append_sheet(wb, ws, "Members");
   // XLSX.writeFile(wb, globalFilePath);
//}



function readGlobal() {
    const filePath = path.join(__dirname, "../global_members.xlsx");
    if (!fs.existsSync(filePath)) {
        return [];
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { defval: "" });

    return data;
}




// Date format dd/mm/yyyy (for welcome)
function getTodayDate() {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    return dd + "/" + mm + "/" + yyyy;
}

// Date format mm/dd/yyyy (for global member record)
function getCurrentDate() {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
}

// Generate Member Code
function generateMemberCode() {
    const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
    let code = "LEN-";
    for (let i = 0; i < 6; i++) {
        code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return code;
}



//part 2
// ---------------------------
// REGISTER
// ---------------------------
router.post("/register", async (req, res) => {

    try {
        const {
            firstName, lastName, spouseName, dob, degree, branch, graduationYear,
            mastersBranch, profession, mobileNumber, EmailAddress, EmployerBusinessName,
            country, state, city, NativeLocation, username, password, recommenderCode
        } = req.body;

        let members = readGlobal();
console.log(members[0]);   // ✅ PUT HERE (temporary)
        // remove blank rows
        members = members.filter(m => m.username && m.username.toString().trim() !== "");

        // allow first member without recommender
        if (members.length > 0) {

            if (!recommenderCode || recommenderCode.trim() === "") {
                return res.json({ message: "Recommender code required." });
            }

            const found = members.find(m => m.memberCode === recommenderCode.trim());
            if (!found) {
                return res.json({ message: "Invalid recommender code." });
            }
        }

        // username duplicate check
        if (members.find(u => u.username === username)) {
            return res.json({ message: "Username already exists." });
        }

        const hashedPassword = await bcrypt.hash(password, 10);
        const memberCode = generateMemberCode();

        const newUser = {
            firstName,
            lastName,
            spouseName,
            dob,
            degree,
            branch,
            graduationYear,
            mastersBranch,
            profession,
            mobileNumber,
            EmailAddress,
            EmployerBusinessName,
            country,
            state,
            city,
            NativeLocation,
            username,
            password: hashedPassword,
            memberCode: memberCode,
            recommenderCode: recommenderCode,
            date: getCurrentDate()
        };

        // save in global file
        members.push(newUser);
        writeGlobal(members);

        // Create personal file first time
        const personalPath = path.join(memberFolderPath, `${firstName}_${lastName}.xlsx`);

        if (!fs.existsSync(personalPath)) {

            const sheetData = [
                [
                    newUser.firstName,
                    newUser.lastName,
                    newUser.username,
                    newUser.EmailAddress,
                    newUser.memberCode
                ],
                [""],
                ["Date", "Activity"]
            ];

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(wb, ws, "Personal");
            XLSX.writeFile(wb, personalPath);
        }

        // AUTO WELCOME MESSAGE INTO get-welcome.xlsx
        const announceFile = path.join(__dirname, "../get-welcome.xlsx");

        let annRows = [];

        if (fs.existsSync(announceFile)) {
            const wb2 = XLSX.readFile(announceFile);
            const ws2 = wb2.Sheets[wb2.SheetNames[0]];
            annRows = XLSX.utils.sheet_to_json(ws2, { header: 1 });
        }

        if (annRows.length === 0) {
            annRows.push(["Date", "Announcement"]);
        }

        const welcomeMessage = "Welcome " + firstName + " " + lastName + " to Alumni Association!";
        annRows.push([getTodayDate(), welcomeMessage]);

        const newWb2 = XLSX.utils.book_new();
        const newWs2 = XLSX.utils.aoa_to_sheet(annRows);
        XLSX.utils.book_append_sheet(newWb2, newWs2, "Welcome");
        XLSX.writeFile(newWb2, announceFile);

        res.json({ message: "Registration successful.", memberCode: memberCode });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Registration error." });
    }

});

// ---------------------------
// LOGIN
// ---------------------------
router.post("/login", async (req, res) => {
    try {
        const { username, password } = req.body;

        let members = readGlobal();
        const user = members.find(u => u.username === username);

        if (!user) return res.json({ message: "User not found." });

        const valid = await bcrypt.compare(password, user.password);
        if (!valid) return res.json({ message: "Invalid password." });

        // Send user data (without password)
        const { password: pw, ...userData } = user;

        res.json({ success: true, user: userData });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Login error." });
    }
});


Part 3

// ---------------------------
// FORGOT PASSWORD
// ---------------------------
router.post("/forgot-password", async (req, res) => {
    try {
        const { username, email, newPassword } = req.body;

        let members = readGlobal();

        const index = members.findIndex(
            u => u.username === username && (u.EmailAddress === email || u.email === email)
        );

        if (index === -1) {
            return res.json({ message: "Invalid username or Email." });
        }

        const hashedPassword = await bcrypt.hash(newPassword, 10);

        members[index].password = hashedPassword;
        members[index].date = getCurrentDate();

        writeGlobal(members);

        res.json({ message: "Password reset successful." });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Reset error." });
    }
});

// ---------------------------
// ADMIN RESET
// ---------------------------
// ---------------------------
// ADMIN RESET
// ---------------------------
router.post("/admin-reset", async (req, res) => {

    const { username, newPassword } = req.body;

    let members = readGlobal();
    const index = members.findIndex(u => u.username === username);

    if (index === -1)
        return res.json({ message: "User not found" });

    const hashed = await bcrypt.hash(newPassword, 10);

    members[index].password = hashed;
    members[index].date = getCurrentDate();

    writeGlobal(members);

    res.json({ message: "Password reset done" });
});


// ---------------------------
// GET USER (FOR editProfile.html)
// ---------------------------
router.post("/get-user", (req, res) => {

    const { username } = req.body;

    let members = readGlobal();
    const user = members.find(u => u.username === username);

    if (!user) {
        return res.json({ success: false, message: "User not found" });
    }

    const { password, ...userData } = user;

    res.json({ success: true, user: userData });
});


// ---------------------------
// UPDATE PROFILE
// ---------------------------
router.post("/update-profile", async (req, res) => {

    try {
        const { username, updatedData } = req.body;

        let members = readGlobal();
        const index = members.findIndex(u => u.username === username);

        if (index === -1) {
            return res.json({ message: "User not found" });
        }

        // update normal fields
        members[index] = {
            ...members[index],
            ...updatedData
        };

        // password update if provided
        if (updatedData.newPassword) {
            const hashed = await bcrypt.hash(updatedData.newPassword, 10);
            members[index].password = hashed;
            delete members[index].newPassword;
        }

        writeGlobal(members);

        res.json({ message: "Profile updated successfully" });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Update error" });
    }
});




// --------------------------- 
// ADMIN MEMBERS LIST (FILTER + SORT)
// ---------------------------

router.get("/admin-members", (req, res) => {

    let sortBy = (req.query.sortBy || "").toLowerCase().trim();
    let filterValue = (req.query.filterValue || "").toLowerCase().trim();

    let members = readGlobal();

    let filtered = members.map(m => ({
    firstName: m.firstName || "",
    lastName: m.lastName || "",
    mobileNumber: m.mobileNumber || "",
    EmailAddress: m.EmailAddress || m.email || "",
    city: (m.city || "").toString().trim(),
    state: (m.state || "").toString().trim(),
    country: (m.country || "").toString().trim(),
    NativeLocation: (m.NativeLocation || "").toString().trim()
}));
    // 🔥 DIRECT FILTER (NO KEY MAP)

   if(filterValue !== ""){
    filtered = filtered.filter(m =>
        (m[sortBy] || "")
            .toString()
            .trim()
            .toLowerCase()
            .includes(filterValue)
    );
}

    res.json({ success: true, members: filtered });
});
