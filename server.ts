import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // Proxy route for data.gouv.fr APIs to bypass CORS CDN caching issues
  app.get("/api/proxy-ads", async (req, res) => {
    try {
      const params = new URLSearchParams();
      for (const key in req.query) {
        params.append(key, String(req.query[key]));
      }
      
      const targetUrl = `https://tabular-api.data.gouv.fr/api/resources/65a9e264-7a20-46a9-9d98-66becb817bc3/data/?${params.toString()}`;
      const apiResponse = await fetch(targetUrl);
      const data = await apiResponse.json();
      
      res.status(apiResponse.status).json(data);
    } catch (err: any) {
      console.error("Proxy error:", err);
      res.status(500).json({ error: err.message });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    // Support Express v4 & v5 wildcard route
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
