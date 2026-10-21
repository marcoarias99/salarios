// Boleta de pago por correo: asunto, cuerpo HTML (con estilos en línea, como lo necesitan Outlook y Gmail)
// y archivo .eml para abrir el correo en Outlook. Funciona en el navegador (window.Boleta) y en Node.
(function (root) {
  'use strict';

  const AZUL = '#2e5c94';
  const AZUL_OSCURO = '#1f3a5f';
  const CELESTE = '#dce6f2';
  const CELESTE_SUAVE = '#eef3f9';
  const BORDE = '#c9d3e0';
  const TEXTO = '#1c2430';

  function escapar(texto) {
    return String(texto == null ? '' : texto)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function colones(n) {
    const texto = Math.abs(n).toLocaleString('en-US', {
      minimumFractionDigits: Number.isInteger(n) ? 0 : 2, maximumFractionDigits: 2,
    });
    return (n < 0 ? '-₡ ' : '₡ ') + texto;
  }

  function fecha(d) {
    if (!(d instanceof Date)) return d || 'N/A';
    return [d.getUTCDate(), d.getUTCMonth() + 1].map(n => String(n).padStart(2, '0')).join('/') + '/' + d.getUTCFullYear();
  }

  function asunto(fila) {
    return `Boleta de Pago - ${fila.nombre}`;
  }

  // ---------- Piezas de la tabla ----------

  const celda = 'padding:8px 10px;border:1px solid ' + BORDE + ';font-size:14px;';
  const etiqueta = (texto, extra) =>
    `<td style="${celda}background:${CELESTE};font-weight:700;${extra || ''}">${escapar(texto)}</td>`;
  const valor = (texto, extra) => `<td style="${celda}${extra || ''}">${escapar(texto)}</td>`;
  const monto = (n, extra) => valor(colones(n), 'text-align:right;white-space:nowrap;' + (extra || ''));
  const titulo = (texto) =>
    `<tr><td colspan="4" style="${celda}background:${AZUL};color:#ffffff;font-weight:700;text-align:center;font-size:13px;">${escapar(texto)}</td></tr>`;
  const subtitulo = (izquierda, derecha) =>
    `<tr><td colspan="2" style="${celda}background:${CELESTE};font-weight:700;text-align:center;">${escapar(izquierda)}</td>` +
    `<td colspan="2" style="${celda}background:${CELESTE};font-weight:700;text-align:center;">${escapar(derecha)}</td></tr>`;

  // Conceptos de pago y deducciones lado a lado; la columna más corta se rellena con celdas vacías.
  function filasEnParalelo(izquierda, derecha) {
    const filas = [];
    for (let i = 0; i < Math.max(izquierda.length, derecha.length); i++) {
      const lado = par => (par ? valor(par[0]) + monto(par[1]) : valor('') + valor(''));
      const fondo = i % 2 ? ` style="background:${CELESTE_SUAVE};"` : '';
      filas.push(`<tr${fondo}>${lado(izquierda[i])}${lado(derecha[i])}</tr>`);
    }
    return filas.join('');
  }

  // ---------- Cuerpo del correo ----------

  // Los montos salen de las mismas columnas del Excel, así que siempre cuadran:
  // Total ingresos − Total deducciones = Salario líquido. Las faltas e incapacidades vienen positivas en el
  // archivo pero restan, así que en la boleta se muestran en negativo.
  function html(fila) {
    const empresa = fila.empresa && fila.empresa !== 'N/A' ? fila.empresa : 'Boleta de Pago';
    const conceptos = [
      ['Salario Bruto:', fila.bruto],
      ['Faltas e incapacidades:', fila.faltas ? -fila.faltas : 0],
      ['Horas extra o feriado:', fila.horasExtra],
    ];
    const deducciones = [
      ['Cargas Laborales:', fila.totalCargas],
      ['Vales o Préstamos:', fila.vales],
      ['Renta:', fila.renta],
    ];
    const cargas = [
      ['Enfermedad y Maternidad:', fila.sem],
      ['Invalidez, vejez y muerte:', fila.ivm],
      ['Banco Popular:', fila.banco],
    ];

    return `<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapar(asunto(fila))}</title>
</head>
<body style="margin:0;padding:24px 12px;background:#f4f6fa;font-family:'Segoe UI',Arial,sans-serif;color:${TEXTO};">
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:600px;margin:0 auto;border-collapse:collapse;background:#ffffff;border:1px solid ${BORDE};">
<tr><td colspan="4" style="padding:18px 16px;background:${AZUL};color:#ffffff;font-size:22px;font-weight:700;letter-spacing:0.5px;">${escapar(empresa)}</td></tr>
<tr>${etiqueta('Boleta de Pago:', 'width:28%;')}<td colspan="3" style="${celda}font-weight:700;color:${AZUL_OSCURO};">${escapar(fila.periodo)}</td></tr>
<tr>${etiqueta('Nombre Completo:')}<td colspan="3" style="${celda}">${escapar(fila.nombre)}</td></tr>
<tr>${etiqueta('Cédula:')}${valor(fila.cedula)}${etiqueta('Salario Líquido:')}${monto(fila.liquido, 'font-weight:700;font-size:15px;')}</tr>
${titulo('Información de Pago Resumida')}
${subtitulo('Conceptos de Pago', 'Deducciones')}
<tr>${etiqueta('Descripción', 'text-align:center;')}${etiqueta('Monto', 'text-align:center;')}${etiqueta('Descripción', 'text-align:center;')}${etiqueta('Monto', 'text-align:center;')}</tr>
${filasEnParalelo(conceptos, deducciones)}
<tr>${etiqueta('Total:')}${monto(fila.totalIngresos, `font-weight:700;background:${CELESTE};`)}${etiqueta('Total:')}${monto(fila.totalDeducciones, `font-weight:700;background:${CELESTE};`)}</tr>
${titulo('Cargas Laborales')}
<tr><td colspan="2" style="${celda}background:${CELESTE};font-weight:700;text-align:center;">Descripción</td><td colspan="2" style="${celda}background:${CELESTE};font-weight:700;text-align:center;">Monto</td></tr>
${cargas.map(([texto, n]) => `<tr><td colspan="2" style="${celda}">${escapar(texto)}</td><td colspan="2" style="${celda}text-align:right;white-space:nowrap;">${escapar(colones(n))}</td></tr>`).join('\n')}
<tr><td colspan="2" style="${celda}background:${CELESTE};font-weight:700;">Total:</td><td colspan="2" style="${celda}background:${CELESTE};font-weight:700;text-align:right;white-space:nowrap;">${escapar(colones(fila.totalCargas))}</td></tr>
<tr>${etiqueta('Fecha de ingreso:')}${valor(fecha(fila.fechaInicio))}${etiqueta('Puesto:')}${valor(fila.puesto)}</tr>
<tr><td colspan="4" style="padding:14px 16px;background:${AZUL_OSCURO};color:#d6deea;font-size:12px;text-align:center;line-height:1.6;">Este documento fue generado automáticamente. Por favor no responder a este correo.<br>Para consultas contacte a su supervisor.</td></tr>
</table>
</body>
</html>`;
  }

  // ---------- Archivo .eml ----------

  function base64(texto) {
    if (typeof Buffer !== 'undefined') return Buffer.from(texto, 'utf8').toString('base64');
    let binario = '';
    new TextEncoder().encode(texto).forEach(b => { binario += String.fromCharCode(b); });
    return btoa(binario);
  }

  const encabezado = texto => `=?UTF-8?B?${base64(texto)}?=`;

  // Correo listo para abrir: "X-Unsent: 1" hace que Outlook lo abra como borrador para enviar.
  function eml(fila) {
    const cuerpo = base64(html(fila)).replace(/.{76}/g, '$&\r\n');
    return [
      `To: ${encabezado(fila.nombre)} <${fila.correo}>`,
      `Subject: ${encabezado(asunto(fila))}`,
      'X-Unsent: 1',
      'MIME-Version: 1.0',
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      cuerpo,
      '',
    ].join('\r\n');
  }

  const api = { asunto, html, eml };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Boleta = api;
})(this);
