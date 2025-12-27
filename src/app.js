import express from "express";
import cors from "cors";
import dotenv from "dotenv";

import authRoutes from "./auth/auth.routes.js";
import companyRoutes from "./companies/company.controller.js";
import driveRoutes from "./drives/drive.controller.js";
import placementRoutes from "./placements/placement.routes.js";
import userRoutes from "./users/user.controller.js";

dotenv.config();
const app = express();

app.use(cors());
app.use(express.json());

app.use("/auth", authRoutes);
app.use("/companies", companyRoutes);
app.use("/drives", driveRoutes);
app.use("/placements", placementRoutes);
app.use("/users", userRoutes);

export default app;
