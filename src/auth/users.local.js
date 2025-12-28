import bcrypt from "bcryptjs";

export const users = [
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
];
