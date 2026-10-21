// Respaldo de un envío: la planilla calculada y el registro de envío (quién, a qué correo, estado y hora),
// con los nombres de carpeta y de archivo que se usan en SharePoint o al descargarlos.
// Funciona en el navegador (window.Respaldo) y en Node.
(function (root) {
  'use strict';

  const P = typeof module !== 'undefined' && module.exports ? require('./planilla.js') : root.Planilla;

  const AZUL = 'FF2E5C94';
  const FORMATO_MONTO = '"₡"#,##0;-"₡"#,##0';

  const dos = n => String(n).padStart(2, '0');

  // "2026-10-Q2": la clave de la quincena (año, mes y número de quincena).
  function codigoQuincena(p) {
    return `${p.anio}-${dos(p.mes + 1)}-Q${p.q}`;
  }

  // Fecha y hora local como texto (ExcelJS guarda las fechas en UTC y saldrían corridas 6 horas).
  function fechaHora(d) {
    return `${dos(d.getDate())}/${dos(d.getMonth() + 1)}/${d.getFullYear()} ${dos(d.getHours())}:${dos(d.getMinutes())}`;
  }

  // Carpeta del respaldo: "2026/2026-10 Segunda quincena de Octubre".
  function carpeta(p) {
    return `${p.anio}/${p.anio}-${dos(p.mes + 1)} ${p.q === 1 ? 'Primera' : 'Segunda'} quincena de ${P.nombreMes(p.mes)}`;
  }

  // La planilla se reemplaza si se vuelve a respaldar; el registro lleva la hora para guardar cada envío.
  function nombres(p, real, momento) {
    const etiqueta = P.etiquetaQuincena(p);
    const sello = `${momento.getFullYear()}-${dos(momento.getMonth() + 1)}-${dos(momento.getDate())} ${dos(momento.getHours())}${dos(momento.getMinutes())}`;
    return {
      planilla: `Planilla - ${etiqueta}.xlsx`,
      registro: `${real ? '' : 'SIMULACION - '}Registro de envío - ${etiqueta} - ${sello}.xlsx`,
    };
  }

  function estadoTexto(item, real) {
    if (item.estado === 'enviado') return real ? 'Enviado' : 'Simulado';
    if (item.estado === 'error') return 'Error';
    return 'No enviado';
  }

  // Una fila por empleado de la planilla, también los que no se enviaron (con el motivo).
  function construirRegistro(ExcelJS, { items, periodo, real, cuenta }) {
    const libro = new ExcelJS.Workbook();
    libro.creator = 'Boletas de pago';
    const hoja = libro.addWorksheet('Registro de envío', { views: [{ state: 'frozen', ySplit: 1 }] });
    hoja.columns = [
      { header: 'Quincena', key: 'quincena', width: 12 },
      { header: 'Periodo', key: 'periodo', width: 38 },
      { header: 'Cédula', key: 'cedula', width: 15 },
      { header: 'Nombre', key: 'nombre', width: 32 },
      { header: 'Correo', key: 'correo', width: 30 },
      { header: 'Empresa', key: 'empresa', width: 16 },
      { header: 'Salario líquido', key: 'liquido', width: 16, style: { numFmt: FORMATO_MONTO } },
      { header: 'Estado', key: 'estado', width: 12 },
      { header: 'Detalle', key: 'detalle', width: 40 },
      { header: 'Fecha y hora', key: 'momento', width: 17 },
      { header: 'Enviado desde', key: 'cuenta', width: 28 },
      { header: 'Modo', key: 'modo', width: 11 },
    ];
    items.forEach(item => {
      const f = item.fila;
      const enviadoONo = item.estado === 'enviado' || item.estado === 'error';
      hoja.addRow({
        quincena: codigoQuincena(periodo),
        periodo: P.textoQuincena(periodo),
        cedula: f.cedula,
        nombre: f.nombre,
        correo: f.correo,
        empresa: f.empresa,
        liquido: f.liquido,
        estado: estadoTexto(item, real),
        detalle: item.estado === 'error' ? item.error : (!enviadoONo ? (item.motivo || 'No se marcó para enviar') : ''),
        momento: item.momento && enviadoONo ? fechaHora(item.momento) : '',
        cuenta: enviadoONo ? (cuenta || '') : '',
        modo: real ? 'Real' : 'Simulado',
      });
    });
    const encabezado = hoja.getRow(1);
    encabezado.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    encabezado.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL } };
    hoja.autoFilter = { from: 'A1', to: 'L1' };
    return libro;
  }

  const api = { codigoQuincena, carpeta, nombres, construirRegistro, fechaHora };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Respaldo = api;
})(this);
