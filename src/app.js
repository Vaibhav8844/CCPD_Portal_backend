import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import session from "express-session";

import authRoutes from "./auth/auth.routes.js";
import companyRoutes from "./companies/company.controller.js";
import driveRoutes from "./drives/drive.controller.js";
import academicYearRoutes from "./routes/academicYear.routes.js";
import placementRoutes from "./placements/placement.routes.js";
import userRoutes from "./users/user.controller.js";
import googleRoutes from "./routes/google.routes.js";
import analyticsRoutes from "./routes/analytics.routes.js";
import enrollStudentsRoutes from "./routes/enrollStudents.routes.js";
import calendarRoutes from "./routes/calendar.routes.js";

dotenv.config();
const app = express();

app.use(cors());
app.use(express.json());

app.use(
  session({
    secret: "google-oauth-secret",
    resave: false,
    saveUninitialized: false,
  })
);

app.use("/auth", authRoutes);
app.use("/auth", googleRoutes);
app.use("/companies", companyRoutes);
app.use("/drives", driveRoutes);
app.use("/academic-year", academicYearRoutes);
app.use("/placements", placementRoutes);
app.use("/users", userRoutes);
app.use("/analytics", analyticsRoutes);
app.use("/data/enroll", enrollStudentsRoutes);
app.use("/calendar", calendarRoutes);

export default app;
