import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import fs from "fs-extra";
import cors from "cors";

// Data storage setup
const DATA_DIR = path.join(process.cwd(), "data");
const USERS_FILE = path.join(DATA_DIR, "users.json");
const STUDENTS_FILE = path.join(DATA_DIR, "students.json");
const ATTENDANCE_FILE = path.join(DATA_DIR, "attendance.json");

async function ensureDataFiles() {
  try {
    await fs.ensureDir(DATA_DIR);
    if (!(await fs.pathExists(USERS_FILE))) await fs.writeJson(USERS_FILE, []);
    if (!(await fs.pathExists(STUDENTS_FILE))) await fs.writeJson(STUDENTS_FILE, []);
    if (!(await fs.pathExists(ATTENDANCE_FILE))) await fs.writeJson(ATTENDANCE_FILE, {});
  } catch (err) {
    console.error("Error ensuring data files:", err);
  }
}

async function startServer() {
  await ensureDataFiles();
  const app = express();
  const PORT = 3000;

  app.use(cors());
  app.use(express.json({ limit: '50mb' })); // Allow large student data

  // --- API Routes ---

  // Health Check
  app.get("/api/health", (req, res) => res.json({ status: "ok" }));

  // User Auth
  app.post("/api/register", async (req, res) => {
    try {
      const { name, username, password } = req.body;
      const users = await fs.readJson(USERS_FILE);
      if (users.find((u: any) => u.username === username)) {
        return res.status(400).json({ error: "Username sudah terdaftar" });
      }
      const newUser = { name, username, password };
      users.push(newUser);
      await fs.writeJson(USERS_FILE, users);
      res.json({ message: "Registrasi berhasil!", user: { name, username } });
    } catch (err) {
      res.status(500).json({ error: "Server error saat registrasi" });
    }
  });

  app.get("/api/users", async (req, res) => {
    try {
      const users = await fs.readJson(USERS_FILE);
      // Only send name and username for security
      res.json(users.map((u: any) => ({ name: u.name, username: u.username })));
    } catch (err) {
      res.json([]);
    }
  });

  app.post("/api/login", async (req, res) => {
    try {
      const { username, password } = req.body;
      const users = await fs.readJson(USERS_FILE);
      const user = users.find((u: any) => u.username === username && u.password === password);
      if (!user) {
        return res.status(401).json({ error: "Username atau password salah" });
      }
      res.json({ message: "Login berhasil!", user: { name: user.name, username: user.username } });
    } catch (err) {
      res.status(500).json({ error: "Server error saat login" });
    }
  });

  // Student Data
  app.get("/api/students", async (req, res) => {
    try {
      const students = await fs.readJson(STUDENTS_FILE);
      res.json(students);
    } catch (err) {
      res.json([]);
    }
  });

  app.post("/api/students/import", async (req, res) => {
    try {
      const { students } = req.body;
      await fs.writeJson(STUDENTS_FILE, students);
      res.json({ message: "Data berhasil diperbarui", count: students.length });
    } catch (err) {
      res.status(500).json({ error: "Gagal menyimpan data siswa" });
    }
  });

  app.delete("/api/students", async (req, res) => {
    try {
      await fs.writeJson(STUDENTS_FILE, []);
      res.json({ message: "Semua data siswa telah dikosongkan" });
    } catch (err) {
      res.status(500).json({ error: "Gagal mengosongkan data" });
    }
  });

  app.delete("/api/students/:nomor", async (req, res) => {
    try {
      const { nomor } = req.params;
      let students = await fs.readJson(STUDENTS_FILE);
      students = students.filter((s: any) => s.nomor !== nomor);
      await fs.writeJson(STUDENTS_FILE, students);
      res.json({ message: "Siswa berhasil dihapus" });
    } catch (err) {
      res.status(500).json({ error: "Gagal menghapus siswa" });
    }
  });

  app.put("/api/students/:oldNomor", async (req, res) => {
    try {
      const { oldNomor } = req.params;
      const updatedStudent = req.body;
      let students = await fs.readJson(STUDENTS_FILE);
      students = students.map((s: any) => s.nomor === oldNomor ? updatedStudent : s);
      await fs.writeJson(STUDENTS_FILE, students);
      res.json({ message: "Data siswa berhasil diupdate" });
    } catch (err) {
      res.status(500).json({ error: "Gagal mengupdate data siswa" });
    }
  });

  // Attendance
  app.get("/api/attendance/:date", async (req, res) => {
    try {
      const { date } = req.params;
      const attendance = await fs.readJson(ATTENDANCE_FILE);
      res.json(attendance[date] || []);
    } catch (err) {
      res.json([]);
    }
  });

  app.post("/api/attendance", async (req, res) => {
    try {
      const { date, data } = req.body;
      const attendance = await fs.readJson(ATTENDANCE_FILE);
      attendance[date] = data;
      await fs.writeJson(ATTENDANCE_FILE, attendance);
      res.json({ message: "Absensi berhasil disimpan" });
    } catch (err) {
      res.status(500).json({ error: "Gagal menyimpan absensi" });
    }
  });

  // Stats
  app.get("/api/stats", async (req, res) => {
    try {
      const students = await fs.readJson(STUDENTS_FILE);
      const attendance = await fs.readJson(ATTENDANCE_FILE);
      const today = new Date().toISOString().split("T")[0];
      const todayAttendance = attendance[today] || [];
      
      res.json({
        totalStudents: students.length,
        attendanceToday: todayAttendance.length,
        historyCount: Object.keys(attendance).length
      });
    } catch (err) {
      res.json({ totalStudents: 0, attendanceToday: 0, historyCount: 0 });
    }
  });

  // --- Vite Handling ---

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server started on http://0.0.0.0:${PORT}`);
  });
}

startServer();
