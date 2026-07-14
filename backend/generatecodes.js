const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

function generateMemberCode(){
    const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
    let code = "LEN-";
    for(let i=0; i<6; i++){
        code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return code;
}

const filePath = path.join(__dirname, "global_members.xlsx");

if(!fs.existsSync(filePath)){
    console.log("global.xlsx not found!");
    process.exit();
}

const wb = XLSX.readFile(filePath);
const ws = wb.Sheets[wb.SheetNames[0]];
let rows = XLSX.utils.sheet_to_json(ws);

// add memberCode if missing
rows = rows.map(m => {
    if(!m.memberCode || m.memberCode.trim() === ""){
        m.memberCode = generateMemberCode();
    }
    return m;
});

const newWs = XLSX.utils.json_to_sheet(rows);
wb.Sheets[wb.SheetNames[0]] = newWs;

XLSX.writeFile(wb, filePath);

console.log("Codes generated successfully!");