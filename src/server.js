import app from "./app.js";
import authRoutes from "./auth/auth.routes.js";

const PORT = process.env.PORT || 5000;

app.use("/auth", authRoutes);

app.listen(PORT, () => {
  console.log(`Backend running on port ${PORT}`);
});
