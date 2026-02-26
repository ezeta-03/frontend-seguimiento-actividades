if (usuario.rol === "Jefatura" || usuario.rol === "Colaborador") {
  const ini = datos.fechaInicio ? new Date(datos.fechaInicio) : null;
  const fin = datos.fechaFin ? new Date(datos.fechaFin) : null;
  if (ini && fin) {
    const cruce = verificarCruceYSugerencias(datos.responsable, ini, fin, null);
    if (cruce.cruce)
      return {
        success: false,
        cruce: true,
        actividad: cruce.actividad,
        horasLibres: cruce.horasLibres,
      };
  }
}

if (usuario.rol === "Jefatura" || usuario.rol === "Colaborador") {
  const ini = cambios.fechaInicio
    ? new Date(cambios.fechaInicio)
    : fechaInicio
      ? new Date(fechaInicio)
      : null;
  const fin = cambios.fechaFin
    ? new Date(cambios.fechaFin)
    : fechaFin
      ? new Date(fechaFin)
      : null;
  const resp = cambios.responsable || fila[4];
  if (ini && fin) {
    const cruce = verificarCruceYSugerencias(resp, ini, fin, id);
    if (cruce.cruce)
      return {
        success: false,
        cruce: true,
        actividad: cruce.actividad,
        horasLibres: cruce.horasLibres,
      };
  }
}
