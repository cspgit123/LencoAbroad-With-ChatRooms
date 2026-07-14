const express = require("express");
const router = express.Router();
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const multer = require("multer");
const bcrypt = require("bcryptjs");




// ----------------------------
// PATHS
// ----------------------------
const membersFilePath = path.join(__dirname, "..", "global_members.xlsx");
const chatRoomsFilePath = path.join(__dirname, "..", "chatrooms.xlsx");
const chatDataFolder = path.join(__dirname, "..", "chatdata");
const uploadFolder = path.join(__dirname, "..", "uploadChats");

const chatMembersFilePath = path.join(__dirname, "..", "chatMembers.xlsx");

// Ensure folders exist
if (!fs.existsSync(chatDataFolder)) fs.mkdirSync(chatDataFolder);
if (!fs.existsSync(uploadFolder)) fs.mkdirSync(uploadFolder);

// ----------------------------
// MULTER CONFIG FOR CHAT MEDIA
// ----------------------------
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadFolder);
  },
  filename: function (req, file, cb) {
    const roomId = req.body.roomId || "unknownRoom";
    const ext = path.extname(file.originalname);
    const unique = Date.now();
    cb(null, `${roomId}_${unique}${ext}`);
  }
});

const upload = multer({ storage });

// ----------------------------
// HELPERS
// ----------------------------
function readExcel(filePath) {
  if (!fs.existsSync(filePath)) return [];
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws);
}

function writeExcel(filePath, rows) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, filePath);
}

function getTodayDate() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function getDateTime() {
  return new Date().toLocaleString();
}

function ensureChatRoomsFile() {
  if (!fs.existsSync(chatRoomsFilePath)) {
    writeExcel(chatRoomsFilePath, []);
  }
}


function ensureChatMembersFile() {
  if (!fs.existsSync(chatMembersFilePath)) {
    writeExcel(chatMembersFilePath, []);
  }
}



function createRoomFile(roomId) {
  const roomFilePath = path.join(chatDataFolder, `${roomId}.xlsx`);
  if (!fs.existsSync(roomFilePath)) {
    writeExcel(roomFilePath, []);
  }
}

function isLoggedIn(req) {
  return req.session && req.session.member;
}

// ----------------------------
// MEMBER LOGIN
// ----------------------------



router.post("/login", (req, res) => {
  const { username, password } = req.body;

  if (!fs.existsSync(membersFilePath)) {
    return res.json({ success: false, message: "Members file not found" });
  }

  const members = readExcel(membersFilePath);

const found = members.find(m =>
  String(m.username || "").trim() === String(username).trim()
);

if (!found) {
  return res.json({ success: false, message: "Invalid username or password" });
}

const ok = bcrypt.compareSync(
  String(password).trim(),
  String(found.password || "").trim()
);

if (!ok) {
  return res.json({ success: false, message: "Invalid username or password" });
}



  /*const found = members.find(m =>
    String(m.username || "").trim() === String(username).trim() &&
    String(m.password || "").trim() === String(password).trim()
  );*/





  if (!found) {
    return res.json({ success: false, message: "Invalid username or password" });
  }

  const firstName = found.firstName || "";
  const lastName = found.lastName || "";



 req.session.regenerate(() => {
  req.session.member = {
username: String(found.username || "").trim(),    firstName,
    lastName
  };
  
 res.json({
    success: true,
    member: {
      username: found.username,
      firstName,
      lastName
    }
  });

});

});


// ----------------------------
// ADMIN LOGIN
// ----------------------------
router.post("/admin-login", (req, res) => {

  const { username, password } = req.body;

  // change these later
  const ADMIN_USER = "admin";
  const ADMIN_PASS = "admin123";

  if (String(username).trim() === ADMIN_USER && String(password).trim() === ADMIN_PASS) {

    req.session.isAdmin = true;

    return res.json({ success: true });
  }

  res.json({ success: false, message: "Invalid admin login" });
});



// ----------------------------
// ADMIN LOGOUT
// ----------------------------
router.get("/admin-logout", (req, res) => {

  req.session.isAdmin = false;

  res.json({ success: true });
});




// ----------------------------
// MEMBER LOGOUT
// ----------------------------
router.get("/logout", (req, res) => {
  req.session.destroy(err => {
    if (err) {
      return res.json({ success: false, message: "Logout failed" });
    }
    res.clearCookie("connect.sid");
    res.json({ success: true });
  });
});




// ----------------------------
// GET CHATROOMS (VISIBLE BEFORE LOGIN ALSO)
// ----------------------------
router.get("/get-chatrooms", (req, res) => {
  ensureChatRoomsFile();
  const rooms = readExcel(chatRoomsFilePath);

  // only active
  const activeRooms = rooms.filter(r => (r.status || "ACTIVE") === "ACTIVE");

  res.json({ success: true, rooms: activeRooms });
});

// ----------------------------
// CREATE CHATROOM (LOGIN REQUIRED)
// ----------------------------
router.post("/create-chatroom", (req, res) => {
  if (!isLoggedIn(req)) {
    return res.json({ success: false, message: "Login required" });
  }

  const { subject } = req.body;

  if (!subject || subject.trim() === "") {
    return res.json({ success: false, message: "Subject required" });
  }

  ensureChatRoomsFile();
  let rooms = readExcel(chatRoomsFilePath);

  const newId = "CR-" + (rooms.length + 1001);

  const creatorName = `${req.session.member.firstName} ${req.session.member.lastName}`.trim();

  const newRoom = {
    roomId: newId,
    subject: subject.trim(),
    creatorUsername: req.session.member.username,
    creatorName: creatorName,
    createdDate: getTodayDate(),
    status: "ACTIVE"
  };

  rooms.push(newRoom);
  writeExcel(chatRoomsFilePath, rooms);

  // Create empty chat file
  createRoomFile(newId);

  res.json({ success: true, room: newRoom });
});


// ----------------------------
// JOIN ROOM (ONE ROOM ONLY)
// ----------------------------

router.post("/join-room/:roomId", (req, res) => {
  if (!isLoggedIn(req)) {
    return res.json({ success: false, message: "Login required" });
  }

  const roomId = req.params.roomId;

  ensureChatRoomsFile();
  ensureChatMembersFile();

  const rooms = readExcel(chatRoomsFilePath);
  const room = rooms.find(r => r.roomId === roomId);

  if (!room) {
    return res.json({ success: false, message: "Room not found" });
  }

  let members = readExcel(chatMembersFilePath);

  // check if member already joined a room


  const currentUser = String(req.session.member.username || "").trim();

const alreadyJoined = members.find(m =>
  String(m.username || "").trim() === currentUser
);


if (alreadyJoined) {
  return res.json({
    success: false,
    message: "You already joined " + alreadyJoined.roomId + ". Leave current room first."
  });
}

  const memberName = `${req.session.member.firstName} ${req.session.member.lastName}`.trim();

  const joinRow = {
    username: String(req.session.member.username || "").trim(),
    memberName: memberName,
    roomId: room.roomId,
    subject: room.subject,
    joinDate: getTodayDate()
  };

  members.push(joinRow);
  writeExcel(chatMembersFilePath, members);

  res.json({ success: true, message: "Joined room successfully", join: joinRow });
});



// ----------------------------
// GET MY JOINED ROOM
// ----------------------------
router.get("/my-joined-room", (req, res) => {
  if (!isLoggedIn(req)) {
    return res.json({ success: false, message: "Login required" });
  }

  ensureChatMembersFile();
  const members = readExcel(chatMembersFilePath);

  const currentUser = String(req.session.member.username || "").trim();

const myJoin = members.find(m =>
  String(m.username || "").trim() === currentUser
);

  if (!myJoin) {
    return res.json({ success: true, joined: false });
  }

  res.json({ success: true, joined: true, room: myJoin });
});




// ----------------------------
// LEAVE ROOM
// ----------------------------
router.post("/leave-room", (req, res) => {
  if (!isLoggedIn(req)) {
    return res.json({ success: false, message: "Login required" });
  }

  ensureChatMembersFile();
  let members = readExcel(chatMembersFilePath);

  const before = members.length;

const currentUser = String(req.session.member.username || "").trim();

members = members.filter(m =>
  String(m.username || "").trim() !== currentUser
);
  if (members.length === before) {
    return res.json({ success: false, message: "You are not joined in any room" });
  }

  writeExcel(chatMembersFilePath, members);

  res.json({ success: true, message: "You left the room successfully" });
});





// ----------------------------
// DEBUG SESSION
// ----------------------------
router.get("/debug-session", (req, res) => {
  res.json({
    member: req.session.member || null,
    isAdmin: req.session.isAdmin || false,
    sessionID: req.sessionID
  });
});



// ----------------------------
// GET ROOM MESSAGES (LOGIN REQUIRED)
// ----------------------------


router.get("/get-messages/:roomId", (req, res) => {
  if (!isLoggedIn(req)) {
    return res.json({ success: false, message: "Login required" });
  }

  const roomId = req.params.roomId;

  ensureChatMembersFile();
  const members = readExcel(chatMembersFilePath);

  const myJoin = members.find(m =>
    m.username === req.session.member.username &&
    m.roomId === roomId
  );

  // admin can view any room
  const isAdmin = req.session && req.session.isAdmin === true;

  if (!myJoin && !isAdmin) {
    return res.json({ success: false, message: "You are not joined in this room" });
  }

  const roomFilePath = path.join(chatDataFolder, `${roomId}.xlsx`);

  if (!fs.existsSync(roomFilePath)) {
    return res.json({ success: true, messages: [] });
  }

  const messages = readExcel(roomFilePath);
  res.json({ success: true, messages });
});







// ----------------------------
// SEND MESSAGE + MEDIA (LOGIN REQUIRED)
// ----------------------------
router.post("/send-message", upload.single("media"), (req, res) => {
  if (!isLoggedIn(req)) {
    return res.json({ success: false, message: "Login required" });
  }

  const { roomId, message } = req.body;

  if (!roomId) {
    return res.json({ success: false, message: "roomId missing" });
  }

  const roomFilePath = path.join(chatDataFolder, `${roomId}.xlsx`);
  if (!fs.existsSync(roomFilePath)) {
    return res.json({ success: false, message: "Room not found" });
  }

  let messagesList = readExcel(roomFilePath);

  let mediaFile = "";
  let mediaType = "";

  if (req.file) {
    mediaFile = req.file.filename;
    mediaType = req.file.mimetype;
  }

  const senderName = `${req.session.member.firstName} ${req.session.member.lastName}`.trim();

  const newMsg = {
    senderName: senderName,
    senderUsername: req.session.member.username,
    message: message || "",
    mediaFile: mediaFile,
    mediaType: mediaType,
    dateTime: getDateTime()
  };

  messagesList.push(newMsg);
  writeExcel(roomFilePath, messagesList);

  res.json({ success: true, msg: newMsg });
});

// ----------------------------
// DELETE CHATROOM (CREATOR OR ADMIN)
// ----------------------------
router.delete("/delete-room/:roomId", (req, res) => {
  const roomId = req.params.roomId;

  ensureChatRoomsFile();
  let rooms = readExcel(chatRoomsFilePath);

  const room = rooms.find(r => r.roomId === roomId);

  if (!room) {
    return res.json({ success: false, message: "Room not found" });
  }

  const isAdmin = req.session && req.session.isAdmin === true;
  const isCreator = req.session.member && req.session.member.username === room.creatorUsername;

  if (!isAdmin && !isCreator) {
    return res.json({ success: false, message: "Not allowed to delete this room" });
  }

  // Remove room from list
  rooms = rooms.filter(r => r.roomId !== roomId);
  writeExcel(chatRoomsFilePath, rooms);

  // Delete chat messages file
  const roomFilePath = path.join(chatDataFolder, `${roomId}.xlsx`);
  if (fs.existsSync(roomFilePath)) {
    fs.unlinkSync(roomFilePath);
  }

  // Delete all uploaded media for this room
  const files = fs.readdirSync(uploadFolder);
  files.forEach(file => {
    if (file.startsWith(roomId + "_")) {
      fs.unlinkSync(path.join(uploadFolder, file));
    }
  });


// Remove all members joined in this room
ensureChatMembersFile();
let joinedMembers = readExcel(chatMembersFilePath);
joinedMembers = joinedMembers.filter(m => String(m.roomId || "").trim() !== roomId);
writeExcel(chatMembersFilePath, joinedMembers);


  res.json({ success: true, message: "Chatroom deleted fully" });
});

// ----------------------------
// DOWNLOAD MEDIA FILES
// ----------------------------
router.get("/download/:filename", (req, res) => {
  const fileName = req.params.filename;
  const filePath = path.join(uploadFolder, fileName);

  if (!fs.existsSync(filePath)) {
    return res.status(404).send("File not found");
  }

  res.download(filePath);
});


// ----------------------------
// GET SINGLE ROOM DETAILS
// ----------------------------
router.get("/get-room/:roomId", (req, res) => {

  const roomId = req.params.roomId;

  ensureChatRoomsFile();
  const rooms = readExcel(chatRoomsFilePath);

  const room = rooms.find(r => r.roomId === roomId);

  if (!room) {
    return res.json({ success: false, message: "Room not found" });
  }

  res.json({ success: true, room });
});


// ----------------------------
// CHECK CURRENT LOGIN STATUS
// ----------------------------
router.get("/whoami", (req, res) => {

  let member = null;
  if (req.session && req.session.member) {
    member = req.session.member;
  }

  let isAdmin = false;
  if (req.session && req.session.isAdmin === true) {
    isAdmin = true;
  }

  res.json({
    success: true,
    member: member,
    isAdmin: isAdmin
  });
});


module.exports = router;