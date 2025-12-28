import express from "express";
import cors from "cors";
import session from "express-session";
import passport from "passport";
import authRoutes from "./auth/auth.routes.js";
import googleAuth from "./auth/google.auth.js";
import driveRoutes from "./routes/drive.routes.js";
import dataRoutes from "./routes/data.routes.js";
import analyticsRoutes from "./routes/analytics.routes.js";
import enrollRoutes from "./routes/enroll.routes.js";
import placementRoutes from "./routes/placement.routes.js";

const app = express();

// CORS configuration for production
const allowedOrigins = process.env.FRONTEND_URL 
  ? [process.env.FRONTEND_URL, 'http://localhost:5173']
  : ['http://localhost:5173'];

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

// Session configuration with environment variable
app.use(
  session({
    secret: process.env.SESSION_SECRET || "fallback-secret-only-for-dev",
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: process.env.NODE_ENV === 'production', // Use HTTPS in production
      httpOnly: true,
      maxAge: 24 * 60 * 60 * 1000 // 24 hours
    }
  })
);

app.use(passport.initialize());
app.use(passport.session());
googleAuth();

// Routes
app.use("/auth", authRoutes);
app.use("/drives", driveRoutes);
app.use("/data", dataRoutes);
app.use("/analytics", analyticsRoutes);
app.use("/enrollment", enrollRoutes);
app.use("/placement", placementRoutes);

// Health check endpoint
app.get("/health", (req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

export default app;