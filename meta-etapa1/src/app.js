import express from "express";
import cors from "cors";
import usersRouter from "./routes/users.routes.js";

const app = express();
app.use(cors());
app.use(express.json());

// rota de saúde
app.get("/health", (req, res) => {
  return res.status(200).json({
    ok: true,
    uptime: process.uptime(),
    env: process.env.NODE_ENV || "development"
  });
});

// rotas de usuários
app.use("/users", usersRouter);

export default app;
