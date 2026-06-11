import express from 'express';
import path from 'path';
import cors from 'cors';
import { createServer as createViteServer } from 'vite';

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(cors());
  app.use(express.json());

  // Use environment variable for password if it exists, otherwise fallback to default
  const validPasswords = process.env.APP_PASSWORD 
    ? [process.env.APP_PASSWORD]
    : ['admin123', 'urbanisme', 'charline'];

  // API endpoints
  app.post('/api/login', (req, res) => {
    const { password } = req.body;
    
    if (validPasswords.includes(password)) {
      res.json({ success: true });
    } else {
      res.status(401).json({ success: false, message: 'Mot de passe incorrect' });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
