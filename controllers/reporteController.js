const reporteService = require('../services/reporteService');

/**
 * Controlador de reportes.
 * Maneja las peticiones HTTP y delega la lógica al servicio.
 */

/**
 * GET / — Obtener todos los reportes
 */
async function getAll(req, res, next) {
  try {
    const reportes = reporteService.obtenerTodos();
    res.json(reportes);
  } catch (error) {
    next(error);
  }
}

/**
 * GET /:id — Obtener un reporte por ID
 */
async function getById(req, res, next) {
  try {
    const reporte = reporteService.obtenerPorId(req.params.id);
    if (!reporte) {
      return res.status(404).json({ error: 'Reporte no encontrado' });
    }
    res.json(reporte);
  } catch (error) {
    next(error);
  }
}

/**
 * POST / — Crear un nuevo reporte
 */
async function create(req, res, next) {
  try {
    const nuevo = reporteService.crear(req.body);
    res.status(201).json(nuevo);
  } catch (error) {
    next(error);
  }
}

/**
 * PUT /:id — Actualizar un reporte existente
 */
async function update(req, res, next) {
  try {
    const actualizado = reporteService.actualizar(req.params.id, req.body);
    if (!actualizado) {
      return res.status(404).json({ error: 'Reporte no encontrado' });
    }
    res.json(actualizado);
  } catch (error) {
    next(error);
  }
}

/**
 * DELETE /:id — Eliminar un reporte
 */
async function remove(req, res, next) {
  try {
    const eliminado = reporteService.eliminar(req.params.id);
    if (!eliminado) {
      return res.status(404).json({ error: 'Reporte no encontrado' });
    }
    res.json({ message: 'Reporte eliminado' });
  } catch (error) {
    next(error);
  }
}

module.exports = {
  getAll,
  getById,
  create,
  update,
  remove,
};
