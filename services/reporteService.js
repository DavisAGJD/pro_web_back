const fs = require('fs');
const path = require('path');
const config = require('../config');
const Reporte = require('../models/Reporte');

/**
 * Servicio de reportes.
 * Maneja toda la lógica de negocio y operaciones de archivo (I/O).
 */

/**
 * Lee todos los reportes del archivo JSON.
 * Crea el directorio y archivo si no existen.
 * @returns {Array} Lista de reportes
 */
function leerReportes() {
  if (!fs.existsSync(config.dataFile)) {
    fs.mkdirSync(config.dataDir, { recursive: true });
    fs.writeFileSync(config.dataFile, '[]');
  }
  return JSON.parse(fs.readFileSync(config.dataFile, 'utf8'));
}

/**
 * Guarda la lista de reportes en el archivo JSON.
 * @param {Array} reportes - Lista de reportes a guardar
 */
function guardarReportes(reportes) {
  fs.writeFileSync(config.dataFile, JSON.stringify(reportes, null, 2));
}

/**
 * Obtiene todos los reportes.
 * @returns {Array} Lista de reportes
 */
function obtenerTodos() {
  return leerReportes();
}

/**
 * Obtiene un reporte por su ID.
 * @param {string} id - ID del reporte
 * @returns {Object|null} Reporte encontrado o null
 */
function obtenerPorId(id) {
  const reportes = leerReportes();
  return reportes.find((r) => r.id === id) || null;
}

/**
 * Crea un nuevo reporte.
 * Valida los datos antes de crear e implementa prevención de duplicados (idempotencia).
 * @param {Object} data - Datos del reporte
 * @returns {Object} Reporte creado o existente
 * @throws {Object} Error de validación con { status, errores }
 */
function crear(data) {
  const { valido, errores } = Reporte.validate(data);
  if (!valido) {
    const error = new Error('Error de validación');
    error.status = 400;
    error.errores = errores;
    throw error;
  }

  const reportes = leerReportes();

  // 1. Evitar duplicar si ya existe el mismo ID (p. ej. sincronización reintentada)
  if (data.id) {
    const existentePorId = reportes.find((r) => r.id === data.id);
    if (existentePorId) {
      return existentePorId;
    }
  }

  // 2. Evitar doble-clic rápido (mismo título, cliente y descripción enviados en un margen de 10 segundos)
  const existentePorContenido = reportes.find((r) =>
    r.titulo === data.titulo.trim() &&
    r.cliente === data.cliente.trim() &&
    r.descripcion === data.descripcion.trim() &&
    Math.abs(new Date(r.fechaCreacion) - new Date(data.fechaCreacion || Date.now())) < 10000
  );
  if (existentePorContenido) {
    return existentePorContenido;
  }

  const nuevo = Reporte.create(data);
  reportes.push(nuevo);
  guardarReportes(reportes);
  return nuevo;
}

/**
 * Actualiza un reporte existente.
 * @param {string} id - ID del reporte a actualizar
 * @param {Object} data - Datos a actualizar
 * @returns {Object|null} Reporte actualizado o null si no se encontró
 */
function actualizar(id, data) {
  const reportes = leerReportes();
  const indice = reportes.findIndex((r) => r.id === id);

  if (indice === -1) {
    return null;
  }

  const { valido, errores } = Reporte.validate({ ...reportes[indice], ...data });
  if (!valido) {
    const error = new Error('Error de validación');
    error.status = 400;
    error.errores = errores;
    throw error;
  }

  reportes[indice] = {
    ...reportes[indice],
    ...data,
    id: reportes[indice].id,
    fechaCreacion: reportes[indice].fechaCreacion,
  };

  guardarReportes(reportes);
  return reportes[indice];
}

/**
 * Elimina un reporte por su ID.
 * @param {string} id - ID del reporte a eliminar
 * @returns {boolean} true si se eliminó, false si no se encontró
 */
function eliminar(id) {
  let reportes = leerReportes();
  const longitudOriginal = reportes.length;
  reportes = reportes.filter((r) => r.id !== id);

  if (reportes.length === longitudOriginal) {
    return false;
  }

  guardarReportes(reportes);
  return true;
}

module.exports = {
  obtenerTodos,
  obtenerPorId,
  crear,
  actualizar,
  eliminar,
};
