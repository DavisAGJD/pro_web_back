const express = require('express');
const cors = require('cors');
const config = require('./config');
const reportesRouter = require('./routes/reportes');
const errorHandler = require('./middleware/errorHandler');

const app = express();

// Middlewares globales
app.use(cors());
app.use(express.json());

// Rutas
app.use('/api/reportes', reportesRouter);

// Endpoint de salud del servidor
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Middleware de manejo de errores (debe ir al final)
app.use(errorHandler);

// Iniciar servidor
app.listen(config.port, () => {
  console.log(`✅ Servidor corriendo en http://localhost:${config.port}`);
});
