import request from "supertest";
import app from "../src/app.js";

describe("ETAPA 2 - Listagem e paginação", () => {
  beforeAll(async () => {
    // cria alguns usuários para testar listagem
    for (let i = 1; i <= 5; i++) {
      await request(app).post("/users").send({
        name: `User${i}`,
        email: `user${i}@example.com`
      });
    }
  });

  test("GET /users sem paginação -> retorna todos", async () => {
    const res = await request(app).get("/users");
    expect(res.status).toBe(200);
    expect(Array.isArray(res.body.data)).toBe(true);
    expect(res.body.total).toBeGreaterThanOrEqual(5);
  });

  test("GET /users?page=1&limit=2 -> retorna 2 usuários", async () => {
    const res = await request(app).get("/users?page=1&limit=2");
    expect(res.status).toBe(200);
    expect(res.body.data.length).toBe(2);
    expect(res.body).toHaveProperty("page", 1);
    expect(res.body).toHaveProperty("limit", 2);
  });

  test("GET /users?page=abc&limit=2 -> 400", async () => {
    const res = await request(app).get("/users?page=abc&limit=2");
    expect(res.status).toBe(400);
    expect(res.body).toHaveProperty("error");
  });
});
