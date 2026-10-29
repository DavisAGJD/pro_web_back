const express = require('express');
const router = express.Router();
const reporteController = require('../controllers/reporteController');

/**
 * Rutas de reportes.
 * Mapeo delgado de verbos HTTP a métodos del controlador.
 */

// Obtener todos los reportes
router.get('/', reporteController.getAll);

// Obtener un reporte por ID
router.get('/:id', reporteController.getById);

// Crear un nuevo reporte
router.post('/', reporteController.create);

// Actualizar un reporte existente
router.put('/:id', reporteController.update);

// Eliminar un reporte
router.delete('/:id', reporteController.remove);

module.exports = router;
