import { v4 as uuidv4 } from "uuid";

// memória simulando banco de dados
const usersById = new Map();
const emailToId = new Map();

function isValidEmail(email) {
  return typeof email === "string" && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

export function createUser({ name, email }) {
  if (typeof name !== "string" || name.trim() === "") {
    const err = new Error("Invalid name");
    err.code = 422;
    throw err;
  }
  if (!isValidEmail(email)) {
    const err = new Error("Invalid email");
    err.code = 422;
    throw err;
  }
  if (emailToId.has(email)) {
    const err = new Error("Email already exists");
    err.code = 409;
    throw err;
  }

  const now = new Date().toISOString();
  const id = uuidv4();
  const user = { id, name: name.trim(), email, createdAt: now };

  usersById.set(id, user);
  emailToId.set(email, id);

  return user;
}

export function getUserById(id) {
  return usersById.get(id) || null;
}

export function listUsers(page, limit) {
  const all = Array.from(usersById.values());

  if (!page || !limit) {
    return all; // lista completa
  }

  const start = (page - 1) * limit;
  const paginated = all.slice(start, start + limit);

  return {
    data: paginated,
    page,
    limit,
    total: all.length
  };
}
