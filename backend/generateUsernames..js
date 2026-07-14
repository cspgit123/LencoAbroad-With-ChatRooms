const XLSX = require("xlsx");
const path = require("path");s
const fs = require("fs");

const globalPath = path.join(__dirname, "global_members.xlsx");

if(!fs.existsSync(globalPath)){
    console.log("global_members.xlsx not found!");
    process.exit();
}

const wb = XLSX.readFile(globalPath);
const ws = wb.Sheets[wb.SheetNames[0]];
let members = XLSX.utils.sheet_to_json(ws);

// helper: clean username text
function clean(str){
    return String(str || "")
        .toLowerCase()
        .trim()
        .replace(/\s+/g, "")
        .replace(/[^a-z0-9]/g, "");
}

let used = new Set();

// store already existing usernames
members.forEach(m=>{
    if(m.username){
        used.add(m.username.toLowerCase());
    }
});

// assign usernames
members.forEach(m => {

    if(!m.firstName || !m.lastName) return;

    if(m.username && String(m.username).trim() !== ""){
        return; // already has username
    }

    let base = clean(m.firstName) + "." + clean(m.lastName);
    let username = base;
    let counter = 2;

    while(used.has(username)){
        username = base + counter;
        counter++;
    }

    m.username = username;
    used.add(username);

});

const newWs = XLSX.utils.json_to_sheet(members);
const newWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWb, newWs, "Members");
XLSX.writeFile(newWb, globalPath);

console.log("✅ Usernames generated and saved in global_members.xlsx");