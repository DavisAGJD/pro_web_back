/**
 * Middleware global de manejo de errores.
 * Captura todos los errores no manejados y devuelve una respuesta JSON consistente.
 */
function errorHandler(err, req, res, _next) {
  // Código de estado: usar el del error si existe, sino 500
  const status = err.status || 500;

  // Construir respuesta de error
  const respuesta = {
    error: err.message || 'Error interno del servidor',
  };

  // Incluir detalle de errores de validación si existen
  if (err.errores && Array.isArray(err.errores)) {
    respuesta.errores = err.errores;
  }

  // Registrar el error en consola (solo en desarrollo)
  if (status === 500) {
    console.error('❌ Error interno:', err);
  }

  res.status(status).json(respuesta);
}

module.exports = errorHandler;
