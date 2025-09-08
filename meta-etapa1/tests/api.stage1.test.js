import request from "supertest";
import app from "../src/app.js";

const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

describe("ETAPA 1 - API Fundamentals", () => {
  test("GET /health -> 200 e shape esperado", async () => {
    const res = await request(app).get("/health");
    expect(res.status).toBe(200);
    expect(res.body).toHaveProperty("ok", true);
    expect(typeof res.body.uptime).toBe("number");
    expect(res.body.uptime).toBeGreaterThan(0);
    expect(typeof res.body.env).toBe("string");
  });

  test("POST /users (válido) -> 201 e retorna usuário com id uuid", async () => {
    const res = await request(app).post("/users").send({
      name: "Victor",
      email: "victor@example.com"
    });
    expect(res.status).toBe(201);
    expect(res.body).toHaveProperty("id");
    expect(uuidRegex.test(res.body.id)).toBe(true);
    expect(res.body).toMatchObject({
      name: "Victor",
      email: "victor@example.com"
    });
    expect(res.body).toHaveProperty("createdAt");
  });

  test("POST /users (email inválido) -> 422", async () => {
    const res = await request(app).post("/users").send({
      name: "Alguém",
      email: "invalido"
    });
    expect(res.status).toBe(422);
    expect(res.body).toHaveProperty("error");
  });

  test("POST /users (name vazio) -> 422", async () => {
    const res = await request(app).post("/users").send({
      name: "",
      email: "a@a.com"
    });
    expect(res.status).toBe(422);
    expect(res.body).toHaveProperty("error");
  });

  test("POST /users (email duplicado) -> 409", async () => {
    const email = "dup@example.com";
    const first = await request(app).post("/users").send({ name: "A", email });
    expect(first.status).toBe(201);

    const second = await request(app).post("/users").send({ name: "B", email });
    expect(second.status).toBe(409);
    expect(second.body).toHaveProperty("error");
  });

  test("GET /users/:id (found/not found)", async () => {
    const created = await request(app).post("/users").send({
      name: "User X",
      email: "x@example.com"
    });
    const id = created.body.id;

    const ok = await request(app).get(`/users/${id}`);
    expect(ok.status).toBe(200);
    expect(ok.body).toHaveProperty("id", id);

    const miss = await request(app).get("/users/00000000-0000-4000-8000-000000000000");
    expect(miss.status).toBe(404);
    expect(miss.body).toHaveProperty("error");
  });
});
