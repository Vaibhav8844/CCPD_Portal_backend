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

// CORS configuration for production
const allowedOrigins = process.env.FRONTEND_URL 
  ? [process.env.FRONTEND_URL, 'http://localhost:5173', 'http://localhost:5174']
  : ['http://localhost:5173', 'http://localhost:5174'];

app.use(cors({
  origin: function (origin, callback) {
    // Allow requests with no origin (like mobile apps or curl requests)
    if (!origin) return callback(null, true);
    if (allowedOrigins.indexOf(origin) === -1) {
      const msg = 'The CORS policy for this site does not allow access from the specified Origin.';
      return callback(new Error(msg), false);
    }
    return callback(null, true);
  },
  credentials: true
}));

app.use(express.json());

app.use(
  session({
    secret: process.env.SESSION_SECRET || "fallback-secret-only-for-dev",
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: process.env.NODE_ENV === 'production',
      httpOnly: true,
      maxAge: 24 * 60 * 60 * 1000
    }
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

// Health check endpoint
app.get("/health", (req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

export default app;
