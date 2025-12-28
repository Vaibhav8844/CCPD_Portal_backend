import bcrypt from "bcryptjs";
import fs from "fs";
import path from "path";

// Default/seed users (always available)
const defaultUsers = [
  {
    id: "1",
    name: "Admin User",
    username: "admin@ccpd.edu",
    passwordHash: bcrypt.hashSync("admin123", 10),
    role: "ADMIN",
  },
  {
    id: "2",
    name: "Calendar Team",
    username: "calendar@ccpd.edu",
    passwordHash: bcrypt.hashSync("calendar123", 10),
    role: "CALENDAR_TEAM",
  },
  {
    id: "3",
    name: "Company SPOC",
    username: "spoc@company.com",
    passwordHash: bcrypt.hashSync("spoc123", 10),
    role: "SPOC",
  },
  {
    id: "4",
    name: "Data Team",
    username: "data@ccpd.edu",
    passwordHash: bcrypt.hashSync("data123", 10),
    role: "DATA_TEAM",
  },
  {
    id: "5",
    name: "Vaibhav Prasad",
    username: "vj22ecb0b31@student.nitw.ac.in",
    passwordHash: bcrypt.hashSync("vaibhav123", 10),
    role: "CALENDAR_TEAM",
  },
];

// Path to persistent users file
const usersFile = path.resolve(process.cwd(), "data", "registered_users.json");

// Load registered users from file
function loadRegisteredUsers() {
  try {
    if (fs.existsSync(usersFile)) {
      return JSON.parse(fs.readFileSync(usersFile, "utf-8"));
    }
  } catch (err) {
    console.error("[users] Failed to load registered users:", err);
  }
  return [];
}

// Save registered users to file
export function saveUsers(registeredUsers) {
  try {
    const dir = path.dirname(usersFile);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(usersFile, JSON.stringify(registeredUsers, null, 2));
    console.log("[users] Saved registered users to file");
  } catch (err) {
    console.error("[users] Failed to save users:", err);
    throw err;
  }
}

// Get only registered users (for saving)
export function getRegisteredUsers() {
  return users.filter((u) => !defaultUsers.find((du) => du.id === u.id));
}

// Merge default users with registered users
const registeredUsers = loadRegisteredUsers();
export let users = [...defaultUsers, ...registeredUsers];
