import { Router } from "express";
import { createUser, getUserById, listUsers } from "../services/users.service.js";

const router = Router();

// POST /users -> cria usuário
router.post("/", (req, res) => {
  try {
    const { name, email } = req.body;
    const user = createUser({ name, email });
    return res.status(201).json(user);
  } catch (err) {
    const status = err.code || 500;
    return res.status(status).json({ error: err.message || "Internal error" });
  }
});

// GET /users/:id -> busca usuário por id
router.get("/:id", (req, res) => {
  const { id } = req.params;
  const user = getUserById(id);
  if (!user) return res.status(404).json({ error: "User not found" });
  return res.status(200).json(user);
});

// GET /users -> lista todos (com paginação opcional)
router.get("/", (req, res) => {
  const { page, limit } = req.query;

  // sem paginação
  if (!page && !limit) {
    const users = listUsers();
    return res.status(200).json({
      data: users,
      total: users.length
    });
  }

  const p = parseInt(page, 10);
  const l = parseInt(limit, 10);

  if (isNaN(p) || isNaN(l) || p <= 0 || l <= 0) {
    return res.status(400).json({ error: "Parâmetros inválidos" });
  }

  const result = listUsers(p, l);
  return res.status(200).json(result);
});

export default router;
