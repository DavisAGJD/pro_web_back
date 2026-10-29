const path = require('path');

/**
 * Configuración centralizada del servidor
 */
const config = {
  // Puerto del servidor
  port: process.env.PORT || 3001,

  // Ruta al archivo de datos JSON
  dataDir: path.join(__dirname, '..', 'data'),
  dataFile: path.join(__dirname, '..', 'data', 'reportes.json'),

  // Valores permitidos para campos enumerados
  prioridadesValidas: ['alta', 'media', 'baja'],
  estadosValidos: ['pendiente', 'en_proceso', 'completado'],

  // Valores por defecto
  defaults: {
    ubicacion: '',
    prioridad: 'media',
    estado: 'pendiente',
    sincronizado: true,
  },
};

module.exports = config;
