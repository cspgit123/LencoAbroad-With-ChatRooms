
const welcomeRoutes = require("./routes/welcomeRoutes");
const newsRoutes = require("./routes/newsRoutes");
const eventsRoutes = require("./routes/eventsRoutes");
const chatRoomsRoutes = require("./routes/chatRoomsRoutes");
const session = require("express-session");

//UPLOAD SETUP//

const express = require("express");
const XLSX = require("xlsx");
const bcrypt = require("bcrypt");
const fs = require("fs");
const path = require("path");
const cors = require("cors");


const multer = require("multer");

// MAIN uploads folder
const uploadsFolder = path.join(__dirname, "../uploads");

// create uploads folder if missing
if (!fs.existsSync(uploadsFolder)) {
    fs.mkdirSync(uploadsFolder);
}

// function to calculate folder size
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


const storage = multer.diskStorage({

destination: function (req, file, cb) {
cb(null, path.join(__dirname, "../uploads"));
},

filename: function (req, file, cb) {
cb(null, Date.now() + "_" + file.originalname);
}
});


const upload = multer({
    storage: storage,
    limits: { fileSize: 20 * 1024 * 1024 }   // 20 MB limit
});





const app = express();

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: "lenco_secret_123",
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 1000 * 60 * 60 * 24 }
}));


app.use("/chat" , chatRoomsRoutes);





//const PORT = 3000;
const PORT = process.env.PORT || 3000;

const ADMIN_USERNAME = "admin";
const ADMIN_PASSWORD = "admin123";   // change later


app.use("/uploads", express.static(path.join(__dirname, "../uploads")));
app.use("/newsEvents", express.static(path.join(__dirname, "newsEvents")));


// API ROUTES
app.use("/chat" , chatRoomsRoutes);
app.use(welcomeRoutes);
app.use(newsRoutes);
app.use(eventsRoutes);

// index pics API
app.get("/get-indexpics", (req, res) => {

    const filePath = path.join(__dirname, "indexPics.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ rows: [] });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    res.json({ rows });
});

// ======================
// STATIC FRONTEND (IMPORTANT)
// ======================
const frontendPath = path.join(__dirname, "../frontend");
//app.use(express.static(frontendPath));
app.use(express.static(path.join(__dirname, "../frontend")));

// ======================
// ROOT PAGE (OPTIONAL)
// ======================
app.get("/", (req, res) => {
    res.sendFile(path.join(frontendPath, "index_option.html"));
});

// ---------------------------
// Paths
// ---------------------------
const globalFilePath = path.join(__dirname, "global_members.xlsx");
const memberFolderPath = path.join(__dirname, "memberfiles");

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

function getTodayDate(){
    const today = new Date();
    const dd = String(today.getDate()).padStart(2,"0");
    const mm = String(today.getMonth()+1).padStart(2,"0");
    const yyyy = today.getFullYear();
    return dd + "/" + mm + "/" + yyyy;
}


function readGlobal() {
    ensureGlobalFile();
    const wb = XLSX.readFile(globalFilePath);
    const ws = wb.Sheets["Members"];
    return XLSX.utils.sheet_to_json(ws);
}


function getUserStorageLimit(username){
    let members = readGlobal();
    const user = members.find(u => u.username === username);

    // default limit if not found
    if(!user) return 200 * 1024 * 1024; // 200 MB

    // read from global_members.xlsx column: storageLimitMB
    const mb = parseInt(user.storageLimitMB);

    if(!mb || mb <= 0){
        return 200 * 1024 * 1024; // default 200MB
    }
    return mb * 1024 * 1024;
}


function writeGlobal(data) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, "Members");
    XLSX.writeFile(wb, globalFilePath);
}

function getCurrentDate() {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
}


// ---------------------------
// ADMIN LOGIN ROUTE
// ---------------------------
  
/////// Add Adnin Route from here To be removed
app.post("/admin-login", (req, res) => {
  const { username, password } = req.body;
 if(username === ADMIN_USERNAME && password === ADMIN_PASSWORD){
  return res.json({ success:true });
  }
   res.json({ success:false, message:"Invalid Admin Login" });
});
//////Added Admin Route upto here.


//Update Dropdown Route

app.get("/get-location-list", (req, res) => {

    const filePath = path.join(__dirname, "global_members.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ countries: [], states: [], cities: [] });





app.get("/get-member-filter-values", (req, res) => {

    const filePath = path.join(__dirname, "global_members.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ countries: [], states: [], cities: [], natives: [] });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws);

    let countries = new Set();
    let states = new Set();
    let cities = new Set();
    let natives = new Set();

    data.forEach(row => {

        if (row.Country) countries.add(row.Country.toString().trim());
        if (row.State) states.add(row.State.toString().trim());
        if (row.City) cities.add(row.City.toString().trim());

        if (row.NativeLocation || row.Native) {
            natives.add((row.NativeLocation || row.Native).toString().trim());
        }

    });

    res.json({
        countries: Array.from(countries).sort(),
        states: Array.from(states).sort(),
        cities: Array.from(cities).sort(),
        natives: Array.from(natives).sort()
    });
});

    }



    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws);

    let countries = new Set();
    let states = new Set();
    let cities = new Set();

   data.forEach(row => {

    let countryValue = row.Country || row.COUNTRY || row.country;
    let stateValue = row.State || row.STATE || row.state;
    let cityValue = row.City || row.CITY || row.city;

    if (countryValue) countries.add(countryValue.toString().trim());
    if (stateValue) states.add(stateValue.toString().trim());
    if (cityValue) cities.add(cityValue.toString().trim());

});

    res.json({
        countries: Array.from(countries).sort(),
        states: Array.from(states).sort(),
        cities: Array.from(cities).sort()
    });

});


// ---------------------------
// REGISTER Route
// ---------------------------   

function generateMemberCode(){
    const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
    let code = "LEN-";
    for(let i=0; i<6; i++){
        code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return code;
}

app.post("/register", async (req, res) => {

    try {

        const {
            firstName, lastName, spouseName, dob, degree, branch, graduationYear,
            mastersBranch, profession, mobileNumber, EmailAddress, EmployerBusinessName,
            country, state, city, NativeLocation, username, password, recommenderCode
        } = req.body;

        let members = readGlobal();

        // ✅ remove blank rows
        members = members.filter(m => m.username && m.username.toString().trim() !== "");

        // ✅ allow first member without recommender
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

        // ✅ AUTO WELCOME MESSAGE INTO get-welcome.xlsx
        const announceFile = path.join(__dirname, "get-welcome.xlsx");

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
// LOGIN (Update Global)

// ---------------------------
app.post("/login", async (req, res) => {
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


// FORGOT PASSWORD
app.post("/forgot-password", async (req, res) => {
    try {
        const { username, EmailAddress, newPassword } = req.body;

        let members = readGlobal();
        const index = members.findIndex(
            u => u.username === username && u.email === EmailAddress
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

// Reset Route

app.post("/admin-reset", async (req, res) => {
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


//LOAD PERSONAL FILE................


app.post("/load-personal", (req, res) => {

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


//SAVE PERSONAL FILE......

app.post("/save-personal", (req, res) => {
    try {
        const { username, rows } = req.body;
 // 🛡️ SAFETY CHECK (ADD HERE)
        if(!rows || rows.length === 0){
            return res.json({ message: "No data received" });
        }

        console.log("USERNAME:", username);
        console.log("ROWS:", rows);

        let members = readGlobal();
        const user = members.find(u => u.username === username);

        if (!user) return res.json({ message: "User not found" });

        const personalPath = path.join(
            memberFolderPath,
            `${user.firstName}_${user.lastName}.xlsx`
        );

       console.log("PATH:", personalPath);

        const newWs = XLSX.utils.aoa_to_sheet(rows);

        // formatting
        newWs["!cols"] = [
            { wch: 15 },   // Date
            { wch: 120 }   // Activity (long text)
        ];

        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, newWs, "Personal");

        XLSX.writeFile(newWb, personalPath);

        console.log("File written successfully");

        res.json({ message: "Personal file saved" });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Save error" });
    }
});

// ---------------------------
// GET USER (FOR editProfile.html)
// ---------------------------
app.post("/get-user", (req, res) => {

    const { username } = req.body;

    let members = readGlobal();

    const user = members.find(u => u.username === username);

    if(!user){
        return res.json({ success:false, message:"User not found" });
    }

    const { password, ...userData } = user;

    res.json({ success:true, user: userData });
});


// ---------------------------
// UPDATE ROUTE

app.post("/update-profile", async (req,res)=>{

    try{
        const { username, updatedData } = req.body;

        let members = readGlobal();
        const index = members.findIndex(u => u.username === username);

        if(index === -1){
            return res.json({ message:"User not found" });
        }

        // update normal fields
        members[index] = {
            ...members[index],
            ...updatedData
        };

        // password update if provided
        if(updatedData.newPassword){
            const hashed = await bcrypt.hash(updatedData.newPassword, 10);
            members[index].password = hashed;
            delete members[index].newPassword;
        }

        writeGlobal(members);

        res.json({ message:"Profile updated successfully" });

    }catch(err){
        console.error(err);
        res.status(500).json({ message:"Update error" });
    }
});

///////////// upto here added new route


app.post("/update", (req, res) => {

    try {

        const { username, updatedData } = req.body;

        let members = readGlobal();
        const index = members.findIndex(u => u.username === username);

        if (index === -1) {
            return res.json({ message: "User not found." });
        }

        members[index] = {
            ...members[index],
            ...updatedData,
            date: getCurrentDate()
        };

        writeGlobal(members);

        const user = members[index];

        const personalPath = path.join(
            memberFolderPath,
            `${user.firstName}_${user.lastName}.xlsx`
        );

        if (fs.existsSync(personalPath)) {

            const wb = XLSX.readFile(personalPath);
            const ws = wb.Sheets["Personal"];

            const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

            let activityRows = [];

            if (rows.length > 3) {
                activityRows = rows.slice(3);
            }

            const newSheetData = [

                [user.firstName, user.lastName, user.username, user.email],
                [],
                ["Date", "Activity"],
                ...activityRows

            ];

            const newWs = XLSX.utils.aoa_to_sheet(newSheetData);

            newWs["!cols"] = [
                { wch: 15 },
                { wch: 60 }
            ];

            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, newWs, "Personal");

            XLSX.writeFile(newWb, personalPath);

        }

        res.json({ message: "Data updated successfully." });

    } catch (err) {

        console.error(err);
        res.status(500).json({ message: "Update error." });

    }

});
 ////////////////////////////////////////////

// ---------------------------
// APPEND
// ---------------------------
app.post("/append", (req, res) => {

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

            // Preserve all activity rows
            if (rows.length >= 4) {
                activityRows = rows.slice(3);
            }
        }

        // Add new row
        activityRows.push([
            getCurrentDate(),
            data
        ]);

        // Rebuild sheet safely
        const newSheetData = [

            [user.firstName, user.lastName, user.username, user.email],
            [],
            ["Date", "Activity"],
            ...activityRows

        ];

        const newWs = XLSX.utils.aoa_to_sheet(newSheetData);

        newWs["!cols"] = [
            { wch: 15 },
            { wch: 40 }

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

// EDIT ROUTE

app.post("/load-personal", (req, res) => {

    const { username } = req.body;

    let members = readGlobal();
    const user = members.find(u => u.username === username);

    if (!user) return res.json({ message: "User not found" });

    const personalPath = path.join(
        memberFolderPath,
        `${user.firstName}_${user.lastName}.xlsx`
    );

    if (!fs.existsSync(personalPath))
        return res.json({ message: "Personal file missing" });

    const wb = XLSX.readFile(personalPath);
    const ws = wb.Sheets["Personal"];

    const rows = XLSX.utils.sheet_to_json(ws, { header: 0 });

    res.json({ rows });

});


//SAVE EDITED FILE ROUTE

app.post("/save-personal", (req, res) => {

    const { username, rows } = req.body;

    let members = readGlobal();
    const user = members.find(u => u.username === username);

    if (!user) return res.json({ message: "User not found" });

    const personalPath = path.join(
        memberFolderPath,
        `${user.firstName}_${user.lastName}.xlsx`
    );

    const newWs = XLSX.utils.aoa_to_sheet(rows);

    newWs["!cols"] = [
        { wch: 20 },
        { wch: 20 }
    ];

    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, newWs, "Personal");

    XLSX.writeFile(newWb, personalPath);

    res.json({ message: "Personal file saved" });

});
//
//UPLOAD ROUTE//


//////////////////////////////////////////
app.post("/upload", upload.single("media"), (req, res) => {

    try {

        const username = req.body.username;

        if(!username){
            return res.json({ success:false, message:"Username missing in upload request" });
        }

        if(!req.file){
            return res.json({ success:false, message:"No file uploaded" });
        }

        // folder where files are stored
        const uploadFolder = path.join(__dirname, "../uploads");

        // total used storage
        const usedBytes = getFolderSize(uploadFolder);

        // user allowed limit
        const maxBytes = getUserStorageLimit(username);

        // size of new uploaded file
        const newFileSize = req.file.size;

        // check if limit exceeded
        if(usedBytes + newFileSize > maxBytes){

            // delete the uploaded file immediately
            fs.unlinkSync(req.file.path);

            return res.json({
                success:false,
                message:"Storage limit exceeded. Please delete old uploads or contact admin."
            });
        }

        res.json({
            success:true,
            message:"Upload successful",
            file:req.file.filename
        });

    } catch (err) {
        console.error(err);
        res.status(500).json({ success:false, message:"Upload error" });
    }

});
//////////////////////
app.get("/admin-members", (req, res) => {

    const sortBy = req.query.sortBy || "city";
    const filterValue = (req.query.filterValue || "").trim().toLowerCase();

    let members = readGlobal();

    // keep only required fields
    let filtered = members.map(m => ({
        firstName: m.firstName || "",
        lastName: m.lastName || "",
        mobileNumber: m.mobileNumber || "",
        EmailAddress: m.EmailAddress || m.email || "",
        city: m.city || "",
        state: m.state || "",
        country: m.country || "",
        NativeLocation: m.NativeLocation || ""
    }));

    // apply filter
    if(filterValue !== ""){
        filtered = filtered.filter(m =>
            (m[sortBy] || "").toLowerCase().includes(filterValue)
        );
    }

    // sort
    filtered.sort((a, b) => {
        return (a[sortBy] || "").localeCompare(b[sortBy] || "");
    });

    res.json({ success: true, members: filtered });
});
/////////////////////////////////////////

app.get("/export-member-register", (req, res) => {

    let members = readGlobal();

    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(members);
    XLSX.utils.book_append_sheet(wb, ws, "Members");

    // Today's date
    const today = new Date();
///
    const now = new Date();

const yyyy = now.getFullYear();
const mm = String(now.getMonth() + 1).padStart(2, "0");
const dd = String(now.getDate()).padStart(2, "0");

const hh = String(now.getHours()).padStart(2, "0");
const mi = String(now.getMinutes()).padStart(2, "0");
const ss = String(now.getSeconds()).padStart(2, "0");

const fileName =
`LencoAbroad_MemberRegister_${yyyy}-${mm}-${dd}_${hh}-${mi}-${ss}.xlsx`;
///
    const exportFile = path.join(__dirname, fileName);

    XLSX.writeFile(wb, exportFile);

    res.download(exportFile, fileName);

});
/////////////////////////////////////////


app.get("/get-member-filter-values", (req, res) => {

    const filePath = path.join(__dirname, "global_members.xlsx");

    if (!fs.existsSync(filePath)) {
        return res.json({ countries: [], states: [], cities: [], natives: [] });
    }

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws);

    let countries = new Set();
    let states = new Set();
    let cities = new Set();
    let natives = new Set();

    data.forEach(row => {

        let countryValue = row.Country || row.COUNTRY || row.country;
        let stateValue = row.State || row.STATE || row.state;
        let cityValue = row.City || row.CITY || row.city;
        let nativeValue = row.NativeLocation || row.NATIVELOCATION || row.nativeLocation || row.Native;

        if (countryValue) countries.add(countryValue.toString().trim());
        if (stateValue) states.add(stateValue.toString().trim());
        if (cityValue) cities.add(cityValue.toString().trim());
        if (nativeValue) natives.add(nativeValue.toString().trim());
    });

    res.json({
        countries: Array.from(countries).sort(),
        states: Array.from(states).sort(),
        cities: Array.from(cities).sort(),
        natives: Array.from(natives).sort()
    });

});




app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});