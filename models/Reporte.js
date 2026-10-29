const config = require('../config');

/**
 * Modelo de datos para Reporte.
 * Define la estructura, validación y creación de reportes.
 */
const Reporte = {
  /**
   * Valida los datos de un reporte.
   * @param {Object} data - Datos a validar
   * @returns {{ valido: boolean, errores: string[] }}
   */
  validate(data) {
    const errores = [];

    // Campos requeridos
    if (!data.titulo || typeof data.titulo !== 'string' || !data.titulo.trim()) {
      errores.push('El campo "titulo" es requerido y debe ser un texto válido');
    }
    if (!data.descripcion || typeof data.descripcion !== 'string' || !data.descripcion.trim()) {
      errores.push('El campo "descripcion" es requerido y debe ser un texto válido');
    }
    if (!data.cliente || typeof data.cliente !== 'string' || !data.cliente.trim()) {
      errores.push('El campo "cliente" es requerido y debe ser un texto válido');
    }

    // Campos opcionales con valores restringidos
    if (data.prioridad && !config.prioridadesValidas.includes(data.prioridad)) {
      errores.push(`El campo "prioridad" debe ser uno de: ${config.prioridadesValidas.join(', ')}`);
    }
    if (data.estado && !config.estadosValidos.includes(data.estado)) {
      errores.push(`El campo "estado" debe ser uno de: ${config.estadosValidos.join(', ')}`);
    }

    return {
      valido: errores.length === 0,
      errores,
    };
  },

  /**
   * Crea un nuevo objeto Reporte preservando el ID del cliente si existe
   * (fundamental para sincronización offline-first PWA sin duplicados).
   * @param {Object} data - Datos del reporte
   * @returns {Object} Reporte completo
   */
  create(data) {
    return {
      id: (data.id && String(data.id).trim()) || (Date.now().toString(36) + Math.random().toString(36).slice(2, 7)),
      titulo: data.titulo.trim(),
      descripcion: data.descripcion.trim(),
      cliente: data.cliente.trim(),
      ubicacion: (data.ubicacion && data.ubicacion.trim()) || config.defaults.ubicacion,
      prioridad: data.prioridad || config.defaults.prioridad,
      estado: data.estado || config.defaults.estado,
      fechaCreacion: data.fechaCreacion || new Date().toISOString(),
      sincronizado: config.defaults.sincronizado,
    };
  },
};

module.exports = Reporte;
