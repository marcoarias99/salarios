// Pantalla principal: subir el Excel, mostrar lo leído, elegir la quincena y descargar la planilla calculada.
(function () {
  'use strict';

  const P = window.Planilla;
  const $ = selector => document.querySelector(selector);

  const estado = {
    archivo: null,
    hoja: null,
    lectura: null,
    opciones: P.opcionesQuincena(new Date()),
    periodo: null,
    tasas: cargarTasas(),
    filas: [],
  };
  estado.periodo = estado.opciones[0];

  // ---------- Formato ----------

  function colones(n) {
    const entero = Number.isInteger(Math.round(n * 100) / 100);
    const texto = Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: entero ? 0 : 2, maximumFractionDigits: 2 });
    return (n < 0 ? '-₡' : '₡') + texto;
  }

  function fecha(d) {
    if (!(d instanceof Date)) return d || '';
    return [d.getUTCDate(), d.getUTCMonth() + 1].map(n => String(n).padStart(2, '0')).join('/') + '/' + d.getUTCFullYear();
  }

  function empleados(n) {
    return `${n} ${n === 1 ? 'empleado' : 'empleados'}`;
  }

  function porcentaje(n) {
    return n.toLocaleString('es-CR', { maximumFractionDigits: 2 }) + '%';
  }

  function crear(etiqueta, atributos, hijos) {
    const el = document.createElement(etiqueta);
    Object.entries(atributos || {}).forEach(([k, v]) => {
      if (v == null || v === false) return;
      if (k === 'texto') el.textContent = v;
      else if (k === 'clase') el.className = v;
      else el.setAttribute(k, v);
    });
    (hijos || []).forEach(h => el.append(h));
    return el;
  }

  function indicador(titulo, valor, detalle, destacado) {
    return crear('div', { clase: 'indicador' + (destacado ? ' destacado' : '') }, [
      crear('span', { texto: titulo }),
      crear('strong', { texto: valor }),
      detalle ? crear('small', { texto: detalle }) : '',
    ]);
  }

  // ---------- Tasas ----------

  // La clave cambió con la plantilla nueva (cargas sobre el total de ingresos, IVM 4,33%), para que no se
  // sigan usando tasas guardadas con la plantilla anterior.
  const CLAVE_TASAS = 'tasas-v2';

  function cargarTasas() {
    try {
      const guardadas = JSON.parse(localStorage.getItem(CLAVE_TASAS));
      if (guardadas && ['sem', 'ivm', 'banco'].every(k => typeof guardadas[k] === 'number')) return guardadas;
    } catch (e) { /* sin almacenamiento: se usan los valores por defecto */ }
    return Object.assign({}, P.TASAS_POR_DEFECTO);
  }

  function guardarTasas(tasas) {
    estado.tasas = tasas;
    try { localStorage.setItem(CLAVE_TASAS, JSON.stringify(tasas)); } catch (e) { /* no persiste, pero se usa igual */ }
    calcular();
  }

  function abrirTasas() {
    $('#tasa-sem').value = estado.tasas.sem;
    $('#tasa-ivm').value = estado.tasas.ivm;
    $('#tasa-banco').value = estado.tasas.banco;
    $('#dialogo-tasas').showModal();
  }

  $('#btn-tasas').addEventListener('click', abrirTasas);
  $('#btn-cancelar').addEventListener('click', () => $('#dialogo-tasas').close());
  $('#btn-restaurar').addEventListener('click', () => {
    $('#tasa-sem').value = P.TASAS_POR_DEFECTO.sem;
    $('#tasa-ivm').value = P.TASAS_POR_DEFECTO.ivm;
    $('#tasa-banco').value = P.TASAS_POR_DEFECTO.banco;
  });
  $('#form-tasas').addEventListener('submit', () => {
    guardarTasas({ sem: Number($('#tasa-sem').value), ivm: Number($('#tasa-ivm').value), banco: Number($('#tasa-banco').value) });
  });

  // ---------- Paso 1: subir archivo ----------

  const zona = $('#zona');
  const entrada = $('#archivo');

  entrada.addEventListener('change', () => { if (entrada.files[0]) abrirArchivo(entrada.files[0]); });
  ['dragenter', 'dragover'].forEach(ev => zona.addEventListener(ev, e => { e.preventDefault(); zona.classList.add('arrastrando'); }));
  ['dragleave', 'drop'].forEach(ev => zona.addEventListener(ev, () => zona.classList.remove('arrastrando')));
  zona.addEventListener('drop', e => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) abrirArchivo(e.dataTransfer.files[0]);
  });
  $('#btn-otro').addEventListener('click', reiniciar);

  function mostrarError(mensaje) {
    $('#error').textContent = mensaje;
    $('#error').hidden = !mensaje;
  }

  function reiniciar() {
    Object.assign(estado, { archivo: null, hoja: null, lectura: null, filas: [] });
    entrada.value = '';
    zona.hidden = false;
    $('#archivo-cargado').hidden = true;
    ['#paso-leido', '#paso-quincena', '#paso-resultado'].forEach(s => { $(s).hidden = true; });
    mostrarError('');
  }

  // CSV: se lee como texto (UTF-8, o Windows-1252 si así lo guardó Excel) y sin convertir nada, para que
  // montos como "? 450,000" y fechas como "01-dic.-17" los interprete la planilla con sus propias reglas.
  function leerCsv(datos) {
    let texto;
    try {
      texto = new TextDecoder('utf-8', { fatal: true }).decode(datos);
    } catch (e) {
      texto = new TextDecoder('windows-1252').decode(datos);
    }
    const libro = XLSX.read(texto, { type: 'string', raw: true });
    libro.esCsv = true;
    return libro;
  }

  async function abrirArchivo(archivo) {
    mostrarError('');
    if (!/\.(xlsx|xlsm|xls|csv)$/i.test(archivo.name)) {
      mostrarError('El archivo debe ser de Excel (.xlsx) o CSV (.csv).');
      return;
    }
    let libro;
    try {
      libro = /\.csv$/i.test(archivo.name)
        ? leerCsv(await archivo.arrayBuffer())
        // cellStyles: necesario para saber qué filas están ocultas en el Excel.
        : XLSX.read(await archivo.arrayBuffer(), { type: 'array', cellStyles: true });
    } catch (e) {
      mostrarError('No se pudo abrir el archivo. Verifique que sea un Excel o CSV válido y que no tenga contraseña, o descargue la plantilla (abajo) y copie los datos en ella.');
      return;
    }
    // Siempre la primera hoja, con los nombres de columna de la plantilla; si algo falta, se dice qué.
    const resultado = P.primeraHoja(XLSX, libro);
    if (resultado.error) {
      mostrarError(`${resultado.error} Descargue la plantilla (abajo) y copie los datos en ella.`);
      return;
    }

    estado.archivo = archivo;
    estado.hoja = resultado.hoja;
    $('#archivo-nombre').textContent = archivo.name;
    $('#archivo-detalle').textContent = `${(archivo.size / 1024).toFixed(0)} KB · se lee la hoja "${resultado.hoja.nombre}"`;
    zona.hidden = true;
    $('#archivo-cargado').hidden = false;
    mostrarHoja();
  }

  // ---------- Paso 2: lo que se leyó ----------

  function mostrarHoja() {
    estado.lectura = P.leerPlanilla(estado.hoja);
    mostrarLectura();
    $('#paso-leido').hidden = false;

    // Sin empleados (por ejemplo, la plantilla vacía) no hay nada que calcular ni enviar.
    const vacia = estado.lectura.empleados.length === 0;
    if (vacia) {
      $('#avisos').querySelectorAll('.alerta.ok').forEach(a => a.remove());
      $('#avisos').prepend(alerta('error', 'Esta hoja no tiene empleados para procesar.',
        ['Si es la plantilla, copie los datos a partir de la fila 4 (la fila 3 es el ejemplo y se ignora).']));
      estado.filas = [];
      estado.envio = null;
    } else {
      mostrarQuincenas();
      calcular();
    }
    $('#paso-quincena').hidden = vacia;
    $('#paso-resultado').hidden = vacia;
  }

  // Lista corta de nombres para los avisos: "Ana, Beto y 3 más".
  function nombres(lista) {
    const vistos = Array.from(new Set(lista.map(e => e.nombre || `fila ${e.filaExcel}`)));
    const primeros = vistos.slice(0, 3).join(', ');
    return vistos.length > 3 ? `${primeros} y ${vistos.length - 3} más` : primeros;
  }

  function alerta(clase, titulo, lineas) {
    return crear('div', { clase: 'alerta ' + clase }, [
      crear('strong', { texto: titulo }),
      lineas && lineas.length ? crear('ul', {}, lineas.map(l => crear('li', { texto: l }))) : '',
    ]);
  }

  function mostrarLectura() {
    const { lectura } = estado;
    const notas = lectura.empleados.flatMap(e => e.notas);
    const ajustes = notas.filter(n => !P.esProblema(n)).length;
    const problemas = notas.filter(P.esProblema).length;
    const encontradas = lectura.columnas.filter(c => c.columna).length;
    const conNota = (prueba) => lectura.empleados.filter(e => e.notas.some(prueba));

    $('#resumen-lectura').replaceChildren(
      indicador('Empleados', String(lectura.empleados.length), 'Filas con datos de empleados', true),
      indicador('Hoja', lectura.hoja, `Encabezados en la fila ${lectura.filaEncabezado}`),
      indicador('Ajustadas', String(ajustes), 'Montos vacíos (→ 0), escritos como texto o datos en N/A'),
      indicador('Problemas', String(problemas), 'Datos vacíos, repetidos o no válidos'),
    );

    const avisos = [];
    const repetidas = conNota(n => n.campo === 'cedula' && n.codigo === 'repetido');
    const correosRepetidos = conNota(n => n.campo === 'correo' && n.codigo === 'repetido');
    const correosMalos = conNota(n => n.codigo === 'formato');
    const vacios = conNota(n => n.tipo === 'vacio');
    const lineas = [];
    if (repetidas.length) lineas.push(`Cédula repetida en ${repetidas.length} filas: ${nombres(repetidas)}.`);
    if (correosRepetidos.length) lineas.push(`Correo repetido en ${correosRepetidos.length} filas: ${nombres(correosRepetidos)}.`);
    if (correosMalos.length) lineas.push(`Correo no válido: ${nombres(correosMalos)}.`);
    const cedulasMalas = conNota(n => n.codigo === 'cedula');
    if (cedulasMalas.length) lineas.push(`Cédula con cantidad de dígitos incorrecta: ${nombres(cedulasMalas)}.`);
    const negativos = conNota(n => n.codigo === 'negativo');
    if (negativos.length) lineas.push(`Montos negativos (deben ser positivos): ${nombres(negativos)}.`);
    const decimales = conNota(n => n.codigo === 'decimal');
    if (decimales.length) lineas.push(`Montos con decimales (deben ser colones enteros): ${nombres(decimales)}.`);
    if (vacios.length) lineas.push(`Falta cédula, nombre o correo en ${vacios.length} ${vacios.length === 1 ? 'empleado' : 'empleados'}: ${nombres(vacios)}.`);
    if (lineas.length) avisos.push(alerta('error', 'Revise estos datos antes de continuar (celdas en rojo):', lineas));

    if (ajustes > 0) {
      const sinDato = (key, texto) => {
        const lista = conNota(n => n.campo === key && n.codigo === 'na');
        return lista.length ? `${texto}: ${nombres(lista)}.` : null;
      };
      const lineasNA = [
        sinDato('fechaInicio', 'Fecha de inicio vacía o no válida (se puso N/A)'),
        sinDato('puesto', 'Puesto vacío (se puso N/A)'),
        sinDato('empresa', 'Empresa vacía (se puso N/A)'),
      ].filter(Boolean);
      avisos.push(alerta('aviso',
        'Las celdas en amarillo se ajustaron: montos vacíos (se completaron con 0), montos escritos como texto o datos que se completaron con N/A. Se puede continuar; pase el cursor sobre ellas para ver el detalle.',
        lineasNA));
    }
    if (lectura.ignoradas.length) {
      avisos.push(alerta('info', lectura.ignoradas.length === 1 ? 'Se ignoró 1 fila que no es de empleados:' : `Se ignoraron ${lectura.ignoradas.length} filas que no son de empleados:`,
        lectura.ignoradas.map(f => `Fila ${f.filaExcel} · ${f.motivo}${f.texto ? ` · "${f.texto}"` : ''}`)));
    }
    if (!lineas.length && !ajustes) {
      avisos.push(alerta('ok', 'Todo en orden: todas las celdas tenían valor y no hay datos repetidos.'));
    }
    $('#avisos').replaceChildren(...avisos);

    const ignoradas = lectura.columnasIgnoradas;
    const recalculadas = ignoradas.filter(c => c.recalculada).length;
    const sinUso = ignoradas.length - recalculadas;
    $('#resumen-columnas').textContent = `Columnas encontradas (${encontradas})` +
      (sinUso ? ` · ignoradas (${sinUso})` : '') + (recalculadas ? ` · calculadas por la página (${recalculadas})` : '');
    $('#lista-columnas').replaceChildren(
      ...lectura.columnas.map(c => crear('li', { clase: c.columna ? '' : 'falta' }, [
        crear('span', { clase: 'letra', texto: c.columna || '✕' }),
        crear('span', { texto: c.campo.encabezado + (c.columna ? '' : ' — no encontrada') }),
      ])),
      ...ignoradas.map(c => crear('li', { clase: 'ignorada',
        title: c.recalculada ? 'La página calcula este valor; no se toma del Excel' : 'Esta columna no se usa' }, [
        crear('span', { clase: 'letra', texto: c.columna }),
        crear('span', { texto: `${c.encabezado} — ${c.recalculada ? 'la página la calcula' : 'se ignora'}` }),
      ])),
    );

    const cabecera = crear('tr', {},
      P.CAMPOS.map(c => crear('th', { clase: c.tipo === 'monto' ? 'num' : '', texto: c.etiqueta })));
    const cuerpo = lectura.empleados.map(e => {
      const celdas = P.CAMPOS.map(c => {
        const notas = e.notas.filter(n => n.campo === c.key);
        let texto = e[c.key];
        if (c.tipo === 'monto') texto = colones(texto);
        if (c.tipo === 'fecha') texto = fecha(texto);
        const clases = [c.tipo === 'monto' ? 'num' : '',
          notas.length ? (notas.some(P.esProblema) ? 'falta' : 'relleno') : ''].join(' ').trim();
        return crear('td', { clase: clases, title: notas.length ? notas.map(n => n.texto).join('\n') : null,
          texto: texto || (notas.length ? '—' : '') });
      });
      return crear('tr', { title: `Fila ${e.filaExcel} del Excel` }, celdas);
    });
    $('#tabla-leida').replaceChildren(crear('thead', {}, [cabecera]), crear('tbody', {}, cuerpo));
  }

  // ---------- Paso 3: quincena ----------

  function mostrarQuincenas() {
    const descripciones = { Actual: 'Quincena actual', Siguiente: 'Quincena siguiente', Anterior: 'Quincena anterior' };
    const tarjetas = estado.opciones.map((op, i) => {
      const radio = crear('input', { type: 'radio', name: 'quincena', value: i });
      radio.checked = op === estado.periodo;
      radio.addEventListener('change', () => { estado.periodo = op; calcular(); });
      return crear('label', { clase: 'opcion' }, [radio, crear('div', {}, [
        crear('strong', { texto: P.etiquetaQuincena(op) }),
        crear('span', { texto: descripciones[op.tipo] }),
      ])]);
    });
    $('#opciones-quincena').replaceChildren(...tarjetas, tarjetaOtraQuincena());
  }

  // Cuarta opción: cualquier quincena (por ejemplo, una que no se envió a tiempo). Se arma como la frase
  // "[Primera] quincena de [Junio] de [2026]", con años desde 3 atrás hasta el siguiente.
  function tarjetaOtraQuincena() {
    const anioActual = new Date().getFullYear();
    const inicial = estado.otraQuincena || estado.opciones[0];
    const radio = crear('input', { type: 'radio', name: 'quincena', value: 'otra' });
    radio.checked = estado.periodo.tipo === 'Otra';
    const selQuincena = crear('select', { 'aria-label': 'Quincena' },
      [crear('option', { value: 1, texto: 'Primera' }), crear('option', { value: 2, texto: 'Segunda' })]);
    const selMes = crear('select', { 'aria-label': 'Mes' },
      P.MESES.map((_, i) => crear('option', { value: i, texto: P.nombreMes(i) })));
    const anios = [];
    for (let a = anioActual - 3; a <= anioActual + 1; a++) anios.push(crear('option', { value: a, texto: String(a) }));
    const selAnio = crear('select', { 'aria-label': 'Año' }, anios);
    selQuincena.value = inicial.q;
    selMes.value = inicial.mes;
    selAnio.value = inicial.anio;

    const elegir = () => {
      estado.otraQuincena = { tipo: 'Otra', q: Number(selQuincena.value), mes: Number(selMes.value), anio: Number(selAnio.value) };
      estado.periodo = estado.otraQuincena;
      radio.checked = true;
      calcular();
    };
    radio.addEventListener('change', elegir);
    [selQuincena, selMes, selAnio].forEach(s => s.addEventListener('change', elegir));
    const tarjeta = crear('div', { clase: 'opcion otra-quincena' }, [radio, crear('div', {}, [
      crear('strong', { texto: 'Otra quincena' }),
      crear('div', { clase: 'frase-quincena' }, [selQuincena, 'quincena de', selMes, 'de', selAnio]),
    ])]);
    tarjeta.addEventListener('click', e => { if (e.target.tagName !== 'SELECT' && !radio.checked) elegir(); });
    return tarjeta;
  }

  // ---------- Paso 4: resultado ----------

  // Mismas columnas y orden que el Excel que se descarga: Cédula … Empresa, Periodo actual y luego las calculadas.
  const COLUMNAS_RESULTADO = P.COLUMNAS_SALIDA;
  const GRUPOS_RESULTADO = P.GRUPOS;
  // Cédula y nombre quedan fijos a la izquierda al desplazarse por la tabla.
  const FIJAS = { cedula: 'fija fija-1', nombre: 'fija fija-2' };

  function calcular() {
    if (!estado.lectura) return;
    estado.filas = P.calcularPlanilla(estado.lectura, estado.periodo, estado.tasas);
    estado.envio = null; // si cambian los datos, la quincena o las tasas, el envío se prepara de nuevo
    const t = P.totales(estado.filas);
    const { sem, ivm, banco } = estado.tasas;

    $('#detalle-resultado').textContent =
      `${P.textoQuincena(estado.periodo)} · Tasas: enfermedad y maternidad ${porcentaje(sem)}, IVM ${porcentaje(ivm)}, Banco Popular ${porcentaje(banco)}`;

    mostrarAyuda();

    const negativos = estado.filas.filter(f => f.liquido <= 0);
    $('#avisos-resultado').replaceChildren(...(negativos.length ? [alerta('error',
      `${negativos.length === 1 ? '1 empleado tiene' : `${negativos.length} empleados tienen`} salario líquido negativo o cero. Revise faltas, vales, renta y el salario bruto:`,
      negativos.map(f => `${f.nombre || `Fila ${f.filaExcel}`}: ${colones(f.liquido)}`))] : []));

    $('#totales').replaceChildren(
      indicador('Salario líquido total', colones(t.general.liquido), empleados(t.general.empleados), true),
      ...t.porEmpresa.map(e => indicador(e.empresa === 'N/A' ? 'Sin empresa (N/A)' : e.empresa, colones(e.liquido), empleados(e.empleados))),
      indicador('Cargas laborales', colones(t.general.totalCargas), 'Total de la planilla'),
    );

    const indice = key => COLUMNAS_RESULTADO.findIndex(c => c.key === key);
    const grupos = crear('tr', { clase: 'grupos' }, GRUPOS_RESULTADO.map(([texto, desde, hasta]) =>
      crear('th', { colspan: indice(hasta) - indice(desde) + 1, texto })));
    const cabecera = crear('tr', {}, COLUMNAS_RESULTADO.map(c =>
      crear('th', { clase: [FIJAS[c.key], c.tipo === 'monto' ? 'num' : ''].join(' ').trim(), texto: c.etiqueta })));

    // Cada fila muestra los mismos colores que el paso 2 y que el Excel: amarillo = ajustado, rojo = revisar.
    const cuerpo = estado.filas.map(f => crear('tr', { title: `Fila ${f.filaExcel} del Excel` }, COLUMNAS_RESULTADO.map(c => {
      const notas = f.notas.filter(n => n.campo === c.key);
      const negativo = c.key === 'liquido' ? f.liquido <= 0 : c.tipo === 'monto' && f[c.key] < 0;
      const clases = [FIJAS[c.key], c.tipo === 'monto' ? 'num' : '', c.key === 'liquido' ? 'liquido' : '',
        negativo ? 'negativo' : '', notas.length ? (notas.some(P.esProblema) ? 'falta' : 'relleno') : ''];
      let texto = f[c.key];
      if (c.tipo === 'monto') texto = colones(texto);
      if (c.tipo === 'fecha') texto = fecha(texto);
      return crear('td', { clase: clases.join(' ').trim(), title: notas.length ? notas.map(n => n.texto).join('\n') : null,
        texto: texto || (notas.length ? '—' : '') });
    })));

    const sumas = COLUMNAS_RESULTADO.map(c => {
      if (c.key === 'cedula') return crear('td', { clase: FIJAS.cedula, texto: 'Total' });
      if (c.key === 'nombre') return crear('td', { clase: FIJAS.nombre, texto: empleados(estado.filas.length) });
      if (c.tipo !== 'monto') return crear('td', {});
      return crear('td', { clase: 'num', texto: colones(estado.filas.reduce((s, f) => s + f[c.key], 0)) });
    });
    $('#tabla-resultado').replaceChildren(
      crear('thead', {}, [grupos, cabecera]), crear('tbody', {}, cuerpo), crear('tfoot', {}, [crear('tr', {}, sumas)]));
  }

  // ---------- Ayuda: cómo se calculan las columnas automáticas ----------

  // Se arma con las tasas que se están usando y, como ejemplo, con el primer empleado de la planilla.
  function mostrarAyuda() {
    const { sem, ivm, banco } = estado.tasas;
    const formulas = [
      ['Total ingresos', 'Salario bruto − Faltas e incapacidades + Horas extra o feriado'],
      ['Enfermedad y maternidad', `Total ingresos × ${porcentaje(sem)}`],
      ['Invalidez, vejez y muerte', `Total ingresos × ${porcentaje(ivm)}`],
      ['Banco Popular', `Total ingresos × ${porcentaje(banco)}`],
      ['Total cargas laborales', 'Enfermedad y maternidad + Invalidez + Banco Popular'],
      ['Total deducciones', 'Total cargas laborales + Vales o préstamos + Renta'],
      ['Salario líquido', 'Total ingresos − Total deducciones'],
    ];
    const partes = [
      crear('h3', { texto: '¿Cómo se calculan las columnas automáticas?' }),
      crear('table', { clase: 'ayuda-tabla' }, [crear('tbody', {}, formulas.map(([nombre, formula]) =>
        crear('tr', {}, [crear('th', { texto: nombre }), crear('td', { texto: `= ${formula}` })])))]),
      crear('p', { clase: 'ayuda-nota', texto: 'Cada carga se redondea a colones enteros antes de sumar. Las tasas se cambian con el botón «Tasas» (arriba a la derecha). El salario líquido del archivo no se usa: siempre se calcula.' }),
    ];
    const f = estado.filas[0];
    if (f) {
      partes.push(crear('p', { clase: 'ayuda-ejemplo-titulo', texto: `Ejemplo con ${f.nombre}:` }), crear('ul', { clase: 'ayuda-ejemplo' }, [
        `Total ingresos = ${colones(f.bruto)} − ${colones(f.faltas)} + ${colones(f.horasExtra)} = ${colones(f.totalIngresos)}`,
        `Cargas = ${colones(f.sem)} + ${colones(f.ivm)} + ${colones(f.banco)} = ${colones(f.totalCargas)}`,
        `Deducciones = ${colones(f.totalCargas)} + ${colones(f.vales)} + ${colones(f.renta)} = ${colones(f.totalDeducciones)}`,
        `Salario líquido = ${colones(f.totalIngresos)} − ${colones(f.totalDeducciones)} = ${colones(f.liquido)}`,
      ].map(linea => crear('li', { texto: linea }))));
    }
    $('#ayuda-panel').replaceChildren(...partes);
  }

  // Se abre al pasar el mouse (CSS) o con clic (para pantallas táctiles); se cierra con Esc o clic afuera.
  function abrirAyuda(abierta) {
    $('#ayuda').classList.toggle('abierta', abierta);
    $('#btn-ayuda').setAttribute('aria-expanded', String(abierta));
  }
  $('#btn-ayuda').addEventListener('click', e => {
    e.stopPropagation();
    abrirAyuda(!$('#ayuda').classList.contains('abierta'));
  });
  $('#ayuda-panel').addEventListener('click', e => e.stopPropagation());
  document.addEventListener('click', () => abrirAyuda(false));
  document.addEventListener('keydown', e => { if (e.key === 'Escape') abrirAyuda(false); });

  // ---------- Plantilla ----------

  $('#btn-plantilla').addEventListener('click', async () => {
    const boton = $('#btn-plantilla');
    boton.disabled = true;
    try {
      const datos = await P.construirPlantilla(ExcelJS).xlsx.writeBuffer();
      descargar('Plantilla planilla quincenal.xlsx', datos);
    } catch (e) {
      console.error(e);
      alert('No se pudo crear la plantilla: ' + e.message);
    } finally {
      boton.disabled = false;
    }
  });

  // ---------- Descargar ----------

  $('#btn-descargar').addEventListener('click', async () => {
    const boton = $('#btn-descargar');
    boton.disabled = true;
    boton.textContent = 'Generando…';
    try {
      const libro = P.construirExcel(ExcelJS, { filas: estado.filas });
      const datos = await libro.xlsx.writeBuffer();
      const blob = new Blob([datos], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const enlace = crear('a', { href: URL.createObjectURL(blob), download: `Planilla - ${P.etiquetaQuincena(estado.periodo)}.xlsx` });
      document.body.append(enlace);
      enlace.click();
      enlace.remove();
      setTimeout(() => URL.revokeObjectURL(enlace.href), 2000);
    } catch (e) {
      console.error(e);
      alert('No se pudo generar el Excel: ' + e.message);
    } finally {
      boton.disabled = false;
      boton.textContent = 'Descargar Excel';
    }
  });

  // ---------- Paso 5: revisar y enviar ----------

  // Envío de correos. "prueba" simula (no sale ningún correo); "real" envía con Microsoft 365 desde la
  // cuenta que inició sesión. El modo real solo se puede elegir con la sesión iniciada.
  const PAUSA_ENTRE_CORREOS_MS = 2100; // Exchange permite unos 30 correos por minuto por buzón
  const esperar = ms => new Promise(r => setTimeout(r, ms));
  const ENVIO = {
    modo: 'prueba',
    async enviar(fila, html) {
      if (this.modo === 'real') {
        await Microsoft.enviarCorreo({ para: fila.correo, nombre: fila.nombre, asunto: Boleta.asunto(fila), html });
        return { simulado: false };
      }
      await esperar(150);
      return { simulado: true };
    },
  };

  // Quién se puede enviar: sin correo válido no se puede; con datos en rojo o líquido ≤ 0 queda sin marcar
  // para que alguien lo revise (se puede marcar a mano).
  function revisarEnvio(f) {
    const correoMalo = !f.correo || f.notas.some(n => n.campo === 'correo' && n.codigo === 'formato');
    if (correoMalo) return { bloqueado: true, motivo: 'Sin correo válido: no se puede enviar' };
    const problema = f.notas.find(P.esProblema);
    if (problema) {
      const columna = P.CAMPOS.find(c => c.key === problema.campo).etiqueta.toLowerCase();
      return { motivo: `Revisar ${columna}: ${problema.texto}` };
    }
    if (f.liquido <= 0) return { motivo: 'Revisar: salario líquido negativo o cero' };
    return {};
  }

  function prepararEnvio() {
    estado.envio = {
      actual: 0,
      enviando: false,
      terminado: false, // se terminó una tanda de envío: se muestra el paso 6
      huboReal: false,  // ya salió al menos un correo real: no se puede volver a la planilla
      revisionLista: { estado: 'sin-revisar' }, // quién ya recibió esta quincena, según la lista de SharePoint
      items: estado.filas.map(f => {
        const revision = revisarEnvio(f);
        return { fila: f, ...revision, marcado: !revision.bloqueado && !revision.motivo, estado: 'pendiente', hora: null, error: null, yaEnviado: null };
      }),
    };
  }

  // ---------- Lista "Envíos de boletas" en SharePoint ----------
  // Una fila por boleta enviada de verdad. Al entrar al paso 5 con la sesión iniciada se revisa quién ya
  // recibió la boleta de esta quincena; esas personas quedan sin marcar (se pueden marcar a mano para reenviar).

  const listaConfigurada = () => { const sp = Microsoft.sharepoint(); return Boolean(sp.sitio && sp.lista); };
  // SharePoint necesita la aprobación de un administrador; mientras no la tenga se envía igual, sin lista
  // y con el respaldo descargado.
  const sinAprobacion = () => Boolean(estado.envio && estado.envio.revisionLista.estado === 'sin-aprobacion');
  const usarLista = () => listaConfigurada() && !sinAprobacion();
  // Se reconoce a la persona por la cédula o por el correo.
  const claveCedula = c => String(c || '').replace(/[^0-9A-Za-z]/g, '').toUpperCase();
  const claveCorreo = c => String(c || '').trim().toLowerCase();

  async function revisarYaEnviados() {
    const envio = estado.envio;
    if (!envio || !listaConfigurada() || Microsoft.estado() !== 'conectado') return;
    if (envio.revisionLista.estado === 'revisando' || envio.enviando) return;
    envio.revisionLista = { estado: 'revisando' };
    mostrarEnvio();
    try {
      const sp = Microsoft.sharepoint();
      const filas = await Microsoft.leerEnvios({ sitio: sp.sitio, lista: sp.lista, quincena: Respaldo.codigoQuincena(estado.periodo) });
      if (estado.envio !== envio) return; // se cambió la planilla mientras tanto
      const ultimo = new Map(); // el envío más reciente de cada persona, por cédula y por correo
      const anotar = (clave, previo) => {
        if (!ultimo.has(clave) || previo.momento > ultimo.get(clave).momento) ultimo.set(clave, previo);
      };
      filas.forEach(r => {
        const previo = { momento: new Date(r.Fecha), por: r.EnviadoPor || '' };
        anotar('c:' + claveCedula(r.Cedula), previo);
        anotar('m:' + claveCorreo(r.Correo), previo);
      });
      ultimo.delete('c:');
      ultimo.delete('m:');
      envio.items.forEach(item => {
        item.yaEnviado = ultimo.get('c:' + claveCedula(item.fila.cedula)) || ultimo.get('m:' + claveCorreo(item.fila.correo)) || null;
        if (item.yaEnviado && item.estado === 'pendiente') item.marcado = false;
      });
      envio.revisionLista = { estado: 'listo' };
    } catch (e) {
      console.error(e);
      if (estado.envio !== envio) return;
      envio.revisionLista = e.codigo === 'sin-aprobacion' ? { estado: 'sin-aprobacion' } : { estado: 'error', error: e.message || 'error desconocido' };
    }
    mostrarEnvio();
  }

  // Aviso del paso 5 con el resultado de la revisión.
  function mostrarRevisionLista() {
    const caja = $('#envio-lista');
    const revision = estado.envio.revisionLista;
    const visible = listaConfigurada() && Microsoft.estado() === 'conectado' && revision.estado !== 'sin-revisar';
    caja.hidden = !visible;
    if (!visible) return;
    const quincena = P.etiquetaQuincena(estado.periodo).toLowerCase();
    if (revision.estado === 'revisando') {
      caja.replaceChildren(alerta('info', `Revisando en SharePoint quién ya recibió la boleta de la ${quincena}…`));
    } else if (revision.estado === 'sin-aprobacion') {
      caja.replaceChildren(crear('div', { clase: 'alerta info' }, [
        crear('strong', { texto: 'SharePoint todavía no está aprobado ' }),
        '(falta el permiso de un administrador de Microsoft 365). Puede enviar igual, pero no se revisa quién ya recibió ' +
          'la boleta, el envío no se anota en la lista y el respaldo se descarga en la computadora.',
      ]));
    } else if (revision.estado === 'error') {
      const boton = crear('button', { type: 'button', clase: 'boton secundario pequeno', texto: 'Revisar de nuevo' });
      boton.disabled = estado.envio.enviando;
      boton.addEventListener('click', revisarYaEnviados);
      caja.replaceChildren(crear('div', { clase: 'alerta error' }, [
        crear('strong', { texto: 'No se pudo revisar en SharePoint quién ya recibió esta boleta. ' }),
        `${revision.error}. Revise antes de enviar para no repetir correos. `,
        boton,
      ]));
    } else {
      const ya = estado.envio.items.filter(i => i.yaEnviado).length;
      caja.replaceChildren(ya
        ? alerta('aviso', `${ya === 1 ? '1 persona ya recibió' : `${ya} personas ya recibieron`} la boleta de la ${quincena} (según SharePoint). ` +
            (ya === 1 ? 'Quedó sin marcar; márquela' : 'Quedaron sin marcar; márquelas') + ' a mano solo si quiere reenviarla.')
        : alerta('ok', `Según SharePoint, nadie ha recibido todavía la boleta de la ${quincena}.`));
    }
  }

  // Agrega a la lista la boleta que se acaba de enviar. Si falla, el correo igual quedó enviado.
  async function registrarEnLista(item) {
    const sp = Microsoft.sharepoint();
    const f = item.fila;
    try {
      await Microsoft.registrarEnvio({
        sitio: sp.sitio,
        lista: sp.lista,
        campos: {
          Title: Respaldo.codigoQuincena(estado.periodo),
          Cedula: String(f.cedula || ''),
          Nombre: f.nombre || '',
          Correo: f.correo || '',
          Liquido: f.liquido,
          Fecha: item.momento.toISOString(),
          EnviadoPor: Microsoft.correo(),
        },
      });
      item.registro = 'ok';
      item.registroError = null;
    } catch (e) {
      console.error(e);
      item.registro = 'error';
      item.registroError = e.message || 'error desconocido';
    }
  }

  function mostrarVista(envio) {
    $('#vista-planilla').hidden = envio;
    $('#vista-envio').hidden = !envio;
    window.scrollTo(0, 0);
  }

  $('#btn-ir-envio').addEventListener('click', () => {
    if (!estado.envio) prepararEnvio();
    mostrarVista(true);
    $('#envio-periodo').textContent = `${P.etiquetaQuincena(estado.periodo)} · ${empleados(estado.filas.length)}`;
    mostrarEnvio();
    if (estado.envio.revisionLista.estado === 'sin-revisar') revisarYaEnviados();
  });
  $('#btn-volver').addEventListener('click', () => mostrarVista(false));

  const hora = d => d.toLocaleTimeString('es-CR', { hour: '2-digit', minute: '2-digit' });
  const pendientesMarcados = () => estado.envio.items.filter(i => i.marcado && i.estado !== 'enviado');

  function textoEstado(item) {
    if (item.estado === 'enviado') return [item.simulado ? `Simulado ${item.hora}` : `Enviado ${item.hora}`, 'ok'];
    if (item.estado === 'enviando') return ['Enviando…', 'enviando'];
    if (item.estado === 'error') return [`Error: ${item.error}`, 'error'];
    if (!item.marcado && item.yaEnviado) return ['Ya enviada', 'ya'];
    if (!item.marcado) return ['No se enviará', 'excluido'];
    return ['Pendiente', 'pendiente'];
  }

  function mostrarEnvio() {
    const { items } = estado.envio;
    const enviados = items.filter(i => i.estado === 'enviado').length;
    const errores = items.filter(i => i.estado === 'error').length;
    const excluidos = items.filter(i => !i.marcado && i.estado !== 'enviado').length;
    const porEnviar = pendientesMarcados().length;
    $('#envio-resumen').replaceChildren(
      indicador('Boletas', String(items.length), 'Empleados en la planilla', true),
      indicador('Por enviar', String(porEnviar), 'Marcadas y pendientes'),
      indicador(ENVIO.modo === 'prueba' ? 'Simuladas' : 'Enviadas', String(enviados), 'Confirmadas'),
      indicador('Con error', String(errores), 'Se pueden reintentar'),
      indicador('No se enviarán', String(excluidos), 'Sin marcar o sin correo válido'),
    );

    // Una vez que sale un correo real ya no se vuelve a la planilla: se termina en el paso 6.
    $('#btn-volver').hidden = estado.envio.enviando || estado.envio.huboReal;
    $('.envio-acciones').hidden = estado.envio.terminado;
    mostrarFinal();
    mostrarRevisionLista();

    const boton = $('#btn-enviar');
    const revisando = estado.envio.revisionLista.estado === 'revisando';
    boton.disabled = estado.envio.enviando || revisando || porEnviar === 0;
    boton.textContent = estado.envio.enviando ? 'Enviando…' : revisando ? 'Revisando SharePoint…' :
      ENVIO.modo === 'prueba'
        ? `Simular envío de ${porEnviar} ${porEnviar === 1 ? 'boleta' : 'boletas'}`
        : `Enviar ${porEnviar} ${porEnviar === 1 ? 'boleta' : 'boletas'} (real)`;
    const marcables = items.filter(i => !i.bloqueado && i.estado !== 'enviado');
    $('#chk-todos').checked = marcables.length > 0 && marcables.every(i => i.marcado);
    $('#chk-todos').disabled = estado.envio.enviando || revisando;

    $('#lista-envio').replaceChildren(...items.map((item, i) => {
      const [texto, clase] = textoEstado(item);
      const casilla = crear('input', { type: 'checkbox', 'aria-label': `Enviar a ${item.fila.nombre}` });
      casilla.checked = item.marcado;
      casilla.disabled = item.bloqueado || item.estado === 'enviado' || estado.envio.enviando ||
        estado.envio.revisionLista.estado === 'revisando'; // la revisión podría desmarcar lo que se marque ahora
      casilla.addEventListener('click', e => e.stopPropagation());
      casilla.addEventListener('change', () => { item.marcado = casilla.checked; mostrarEnvio(); });
      const li = crear('li', { clase: 'item-envio' + (i === estado.envio.actual ? ' actual' : '') }, [
        casilla,
        crear('div', { clase: 'item-datos' }, [
          crear('strong', { texto: item.fila.nombre || `Fila ${item.fila.filaExcel}` }),
          crear('span', { texto: item.fila.correo || 'Sin correo' }),
          item.motivo ? crear('span', { clase: 'item-motivo', texto: item.motivo }) : '',
          item.yaEnviado ? crear('span', { clase: 'item-ya', texto:
            `Ya enviada el ${Respaldo.fechaHora(item.yaEnviado.momento)}${item.yaEnviado.por ? ` por ${item.yaEnviado.por}` : ''}` }) : '',
        ]),
        crear('span', { clase: 'estado-envio ' + clase, texto }),
      ]);
      li.addEventListener('click', () => { estado.envio.actual = i; mostrarEnvio(); });
      return li;
    }));
    mostrarPrevia();
  }

  function mostrarPrevia() {
    const { items, actual } = estado.envio;
    const fila = items[actual].fila;
    $('#previa-para').textContent = fila.correo ? `${fila.nombre} <${fila.correo}>` : `${fila.nombre} (sin correo)`;
    $('#previa-asunto').textContent = Boleta.asunto(fila);
    $('#previa-posicion').textContent = `${actual + 1} de ${items.length}`;
    $('#btn-anterior').disabled = actual === 0;
    $('#btn-siguiente').disabled = actual === items.length - 1;
    const marco = $('#previa');
    const html = Boleta.html(fila);
    if (marco.dataset.fila === String(actual) && marco.srcdoc === html) return; // misma boleta: no se recarga
    marco.dataset.fila = String(actual);
    marco.onload = () => { marco.style.height = marco.contentDocument.documentElement.scrollHeight + 'px'; };
    marco.srcdoc = html;
  }

  $('#btn-anterior').addEventListener('click', () => { estado.envio.actual--; mostrarEnvio(); });
  $('#btn-siguiente').addEventListener('click', () => { estado.envio.actual++; mostrarEnvio(); });

  $('#chk-todos').addEventListener('change', e => {
    estado.envio.items.filter(i => !i.bloqueado && i.estado !== 'enviado').forEach(i => { i.marcado = e.target.checked; });
    mostrarEnvio();
  });

  // Un correo como archivo: al abrirlo, Outlook lo muestra como borrador listo para enviar.
  $('#btn-eml').addEventListener('click', () => {
    const fila = estado.envio.items[estado.envio.actual].fila;
    const blob = new Blob([Boleta.eml(fila)], { type: 'message/rfc822' });
    const enlace = crear('a', { href: URL.createObjectURL(blob), download: `${Boleta.asunto(fila)}.eml` });
    document.body.append(enlace);
    enlace.click();
    enlace.remove();
    setTimeout(() => URL.revokeObjectURL(enlace.href), 2000);
  });

  // Advertencias para la confirmación del envío real.
  function avisoReenvio(lista) {
    let texto = '';
    const reenvios = lista.filter(i => i.yaEnviado).length;
    if (reenvios) texto += ` ATENCIÓN: ${reenvios === 1 ? '1 de ellas ya se había enviado' : `${reenvios} de ellas ya se habían enviado`} antes y se enviará de nuevo.`;
    if (listaConfigurada() && ['error', 'sin-aprobacion'].includes(estado.envio.revisionLista.estado)) texto += ' No se pudo revisar en SharePoint quién ya la recibió.';
    return texto;
  }

  // Envía una por una y confirma cada correo en la lista (enviado con la hora, o el error para reintentar).
  // En el envío real cada correo enviado se anota en la lista de SharePoint y al final se guarda el respaldo.
  async function enviarMarcados() {
    if (estado.envio.enviando) return; // un segundo clic no arranca otra tanda
    const lista = pendientesMarcados();
    if (!lista.length) return;
    const prueba = ENVIO.modo === 'prueba';
    const pregunta = prueba
      ? `Se simulará el envío de ${lista.length} boletas de ${P.etiquetaQuincena(estado.periodo)}. No se enviará ningún correo. ¿Continuar?`
      : `ENVÍO REAL: se ${lista.length === 1 ? 'enviará 1 correo' : `enviarán ${lista.length} correos`} a los empleados, desde ${Microsoft.correo()}, ` +
        `con las boletas de ${P.etiquetaQuincena(estado.periodo)}.` + avisoReenvio(lista) + ' Esto no se puede deshacer. ¿Continuar?';
    if (!confirm(pregunta)) return;

    estado.envio.enviando = true;
    estado.envio.terminado = false;
    estado.envio.detenido = null;
    if (!prueba) estado.envio.huboReal = true;
    mostrarConexion();
    let ultimo = 0; // cuándo empezó el correo anterior: la pausa descuenta lo que ya tardó
    for (const item of lista) {
      const falta = PAUSA_ENTRE_CORREOS_MS - (Date.now() - ultimo);
      if (!prueba && ultimo && falta > 0) await esperar(falta);
      ultimo = Date.now();
      item.estado = 'enviando';
      mostrarEnvio();
      try {
        const resultado = await ENVIO.enviar(item.fila, Boleta.html(item.fila));
        item.estado = 'enviado';
        item.simulado = resultado.simulado;
        item.momento = new Date();
        item.hora = hora(item.momento);
        item.error = null;
      } catch (e) {
        item.estado = 'error';
        item.momento = new Date();
        item.error = e.message || 'no se pudo enviar';
        item.incierto = Boolean(e.incierto);
        // Sin sesión o sin permiso fallarían todos los siguientes: se detiene y quedan pendientes.
        if (e.detener) { estado.envio.detenido = item.error; break; }
        continue;
      }
      item.incierto = false;
      if (!prueba && usarLista()) await registrarEnLista(item);
    }
    estado.envio.enviando = false;
    estado.envio.terminado = true;
    $('#final-respaldo').replaceChildren();
    estado.envio.finalizado = { real: !prueba, hora: hora(new Date()), cuenta: Microsoft.correo() };
    mostrarConexion();
    mostrarEnvio();
    window.scrollTo(0, 0);
    if (!prueba && Microsoft.sharepoint().sitio && Microsoft.estado() === 'conectado' && !sinAprobacion()) crearRespaldo();
  }

  $('#btn-enviar').addEventListener('click', enviarMarcados);

  // ---------- Paso 6: envío terminado ----------

  function mostrarFinal() {
    const envio = estado.envio;
    const seccion = $('#paso-final');
    seccion.hidden = !envio || !envio.terminado;
    if (seccion.hidden) return;

    const { real, hora: horaFin, cuenta } = envio.finalizado;
    const enviados = envio.items.filter(i => i.estado === 'enviado' && i.simulado === !real);
    const errores = envio.items.filter(i => i.estado === 'error');
    const noEnviados = envio.items.filter(i => i.estado !== 'enviado' && i.estado !== 'error');

    $('#final-titulo').textContent = real ? 'Envío terminado' : 'Simulación terminada';
    $('#final-detalle').textContent = real
      ? `${P.etiquetaQuincena(estado.periodo)} · enviado desde ${cuenta} · terminó a las ${horaFin}`
      : `${P.etiquetaQuincena(estado.periodo)} · terminó a las ${horaFin} · no se envió ningún correo`;
    $('#final-resumen').replaceChildren(
      indicador(real ? 'Enviadas' : 'Simuladas', String(enviados.length), real ? 'Confirmadas por Microsoft' : 'No salió ningún correo', true),
      indicador('Con error', String(errores.length), errores.length ? 'Se pueden reintentar' : 'Ninguno'),
      indicador('No se enviaron', String(noEnviados.length), 'Sin marcar o sin correo válido'),
    );

    const dudosas = errores.filter(i => i.incierto);
    const paraReintentar = reintentables();
    let aviso;
    if (envio.detenido) {
      const quedan = envio.items.filter(i => i.marcado && i.estado === 'pendiente').length;
      aviso = alerta('error', `Se detuvo el envío: ${envio.detenido}. ${quedan ? `Quedan ${quedan} sin enviar. ` : ''}` +
        'Inicie sesión de nuevo (botón arriba a la derecha) y use "Reintentar".', errores.map(i => `${i.fila.nombre}: ${i.error}`));
    } else if (errores.length) {
      aviso = alerta('error', `${errores.length === 1 ? '1 boleta no se pudo enviar' : `${errores.length} boletas no se pudieron enviar`}:`,
        errores.map(i => `${i.fila.nombre}: ${i.error}`));
    } else if (real) {
      aviso = alerta('ok', `Correos: todas las boletas marcadas se enviaron. Las copias están en Elementos enviados de ${cuenta}.`);
    } else {
      aviso = alerta('info', Microsoft.estado() === 'conectado'
        ? 'Para enviarlas de verdad, elija Modo de envío → Real y vuelva a enviar.'
        : 'Para enviarlas de verdad, inicie sesión con Microsoft y elija Modo de envío → Real.');
    }
    $('#final-aviso').replaceChildren(aviso);
    mostrarRegistroFinal(real ? envio.items.filter(i => i.estado === 'enviado' && !i.simulado && i.registro) : []);
    $('#btn-reintentar').hidden = !paraReintentar.length;
    $('#btn-reintentar').textContent = `Reintentar ${paraReintentar.length === 1 ? '1 boleta' : `${paraReintentar.length} boletas`}`;
    $('#btn-dudosas').hidden = !dudosas.length;
    $('#btn-dudosas').textContent = `Reenviar ${dudosas.length === 1 ? 'la dudosa' : `las ${dudosas.length} dudosas`}`;
  }

  // Confirmación de que cada boleta enviada quedó anotada en la lista de SharePoint.
  function mostrarRegistroFinal(enviados) {
    const caja = $('#final-registro');
    const huboReales = estado.envio.items.some(i => i.estado === 'enviado' && !i.simulado);
    if (huboReales && listaConfigurada() && sinAprobacion()) {
      caja.replaceChildren(alerta('info', 'Registro: SharePoint todavía no está aprobado, así que este envío no quedó anotado en la lista. ' +
        'El registro de envío del respaldo sí lo guarda.'));
      return;
    }
    if (!enviados.length) { caja.replaceChildren(); return; }
    const fallidos = enviados.filter(i => i.registro === 'error');
    const lista = Microsoft.sharepoint().lista;
    if (!fallidos.length) {
      caja.replaceChildren(alerta('ok', `Registro: ${enviados.length === 1 ? 'la boleta enviada quedó anotada' : `las ${enviados.length} boletas enviadas quedaron anotadas`} en la lista "${lista}" de SharePoint.`));
      return;
    }
    const boton = crear('button', { type: 'button', clase: 'boton secundario pequeno', texto: 'Anotar de nuevo' });
    boton.addEventListener('click', async () => {
      boton.disabled = true;
      boton.textContent = 'Anotando…';
      for (const item of fallidos) await registrarEnLista(item);
      mostrarFinal();
    });
    caja.replaceChildren(crear('div', { clase: 'alerta error' }, [
      crear('strong', { texto: `Registro: ${fallidos.length === 1 ? '1 boleta enviada no se pudo anotar' : `${fallidos.length} boletas enviadas no se pudieron anotar`} en la lista "${lista}". ` }),
      'Los correos sí se enviaron. ',
      boton,
      crear('ul', {}, fallidos.map(i => crear('li', { texto: `${i.fila.nombre}: ${i.registroError}` }))),
    ]));
  }

  // ---------- Respaldo (paso 6) ----------

  // Guarda la planilla calculada y el registro de envío. Con SharePoint configurado y la sesión iniciada se
  // suben al sitio (una carpeta por año y quincena); si no, se descargan en la computadora.
  function descargar(nombre, datos) {
    const blob = new Blob([datos], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const enlace = crear('a', { href: URL.createObjectURL(blob), download: nombre });
    document.body.append(enlace);
    enlace.click();
    enlace.remove();
    setTimeout(() => URL.revokeObjectURL(enlace.href), 2000);
  }

  // Se llama sola al terminar un envío real (con SharePoint configurado) y con el botón "Crear respaldo".
  async function crearRespaldo() {
    const envio = estado.envio;
    const real = envio.finalizado.real;
    const momento = new Date();
    const nombres = Respaldo.nombres(estado.periodo, real, momento);
    const boton = $('#btn-respaldo');
    const resultado = $('#final-respaldo');
    boton.disabled = true;
    boton.textContent = 'Creando respaldo…';
    resultado.replaceChildren(alerta('info', 'Respaldo: guardando la planilla y el registro de envío…'));
    try {
      const planilla = await P.construirExcel(ExcelJS, { filas: estado.filas }).xlsx.writeBuffer();
      const registro = await Respaldo.construirRegistro(ExcelJS, {
        items: envio.items, periodo: estado.periodo, real, cuenta: envio.finalizado.cuenta,
      }).xlsx.writeBuffer();

      const sharepoint = Microsoft.sharepoint();
      if (sharepoint.sitio && Microsoft.estado() === 'conectado' && !sinAprobacion()) {
        try {
          const carpeta = `${(sharepoint.carpeta || 'Boletas de pago').replace(/\/+$/, '')}/${Respaldo.carpeta(estado.periodo)}`;
          const tipo = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
          await Microsoft.subirArchivo({ sitio: sharepoint.sitio, ruta: `${carpeta}/${nombres.planilla}`, contenido: planilla, tipo });
          const urlRegistro = await Microsoft.subirArchivo({ sitio: sharepoint.sitio, ruta: `${carpeta}/${nombres.registro}`, contenido: registro, tipo });
          const urlCarpeta = urlRegistro.slice(0, urlRegistro.lastIndexOf('/'));
          resultado.replaceChildren(crear('div', { clase: 'alerta ok' }, [
            crear('strong', { texto: `Respaldo: guardado en SharePoint (${hora(momento)}). ` }),
            crear('a', { href: urlCarpeta, target: '_blank', rel: 'noopener', texto: `abrir la carpeta ${carpeta}` }),
            crear('ul', {}, [crear('li', { texto: nombres.planilla }), crear('li', { texto: nombres.registro })]),
          ]));
          return;
        } catch (e) {
          if (e.codigo !== 'sin-aprobacion') throw e;
          envio.revisionLista = { estado: 'sin-aprobacion' }; // se descarga en su lugar
        }
      }
      descargar(nombres.planilla, planilla);
      await new Promise(r => setTimeout(r, 400)); // algunos navegadores ignoran dos descargas seguidas
      descargar(nombres.registro, registro);
      const motivo = !sharepoint.sitio ? 'SharePoint todavía no está configurado (js/config.js).'
        : Microsoft.estado() !== 'conectado' ? 'Inicie sesión con Microsoft para guardarlo en SharePoint.'
        : 'SharePoint todavía no está aprobado (falta el permiso de un administrador de Microsoft 365).';
      resultado.replaceChildren(alerta('info', `Respaldo descargado en esta computadora (carpeta Descargas). ${motivo}`,
        [nombres.planilla, nombres.registro]));
    } catch (e) {
      console.error(e);
      resultado.replaceChildren(alerta('error', `Respaldo: no se pudo guardar (${e.message}). Use "Crear respaldo" para intentarlo de nuevo.`));
    } finally {
      boton.disabled = false;
      boton.textContent = 'Crear respaldo';
    }
  }

  $('#btn-respaldo').addEventListener('click', crearRespaldo);

  // Lo que se puede reintentar sin riesgo de duplicar: errores claros (Microsoft no aceptó el correo) y las
  // marcadas que quedaron pendientes porque el envío se detuvo. Las dudosas tienen su propio botón.
  function reintentables() {
    return estado.envio.items.filter(i => (i.estado === 'error' && !i.incierto) || (i.marcado && i.estado === 'pendiente'));
  }

  function reintentar(lista) {
    if (estado.envio.enviando) return;
    // Solo se envía lo de esta lista: todo lo demás que no salió queda sin marcar (también los errores).
    estado.envio.items.forEach(i => { if (i.estado !== 'enviado') i.marcado = false; });
    lista.forEach(i => { i.estado = 'pendiente'; i.marcado = true; i.error = null; });
    enviarMarcados();
  }

  $('#btn-reintentar').addEventListener('click', () => reintentar(reintentables()));

  // Dudosas: Microsoft no respondió, así que el correo pudo haber salido. Se reenvían solo si se confirma.
  $('#btn-dudosas').addEventListener('click', () => {
    const dudosas = estado.envio.items.filter(i => i.estado === 'error' && i.incierto);
    if (!confirm(`${dudosas.length === 1 ? 'Esta boleta pudo' : `Estas ${dudosas.length} boletas pudieron`} haber salido aunque Microsoft no respondió: ` +
      `${dudosas.map(i => i.fila.nombre).join(', ')}.\n\nRevise primero Elementos enviados de ${Microsoft.correo()}. ` +
      '¿Ya lo revisó y quiere reenviarlas?')) return;
    reintentar(dudosas);
  });

  // Cierra este envío y vuelve al paso 1 vacío. El modo vuelve a "Simulado" para que nada salga por accidente.
  $('#btn-nuevo').addEventListener('click', () => {
    const pregunta = estado.envio.huboReal
      ? '¿Terminar este envío y empezar uno nuevo? Se cerrará la planilla cargada (lo enviado ya quedó enviado).'
      : '¿Empezar de nuevo? Se cerrará la planilla cargada.';
    if (!confirm(pregunta)) return;
    estado.envio = null;
    ENVIO.modo = 'prueba';
    $('#prueba-resultado').textContent = '';
    reiniciar();
    mostrarVista(false);
    mostrarConexion();
  });

  // ---------- Microsoft 365: sesión, modo de envío y correo de prueba ----------

  const CORREO_VALIDO = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  // Botón del encabezado, aviso del paso 5 y controles de envío según el estado de la conexión.
  function mostrarConexion() {
    const conexion = Microsoft.estado();
    const conectado = conexion === 'conectado';
    const boton = $('#btn-microsoft');
    boton.classList.toggle('conectado', conectado);
    boton.textContent = conectado ? `${Microsoft.correo()} · Cerrar sesión` : 'Iniciar sesión con Microsoft';
    boton.title = {
      'sin-configurar': 'Falta el Client ID y el Tenant ID en js/config.js',
      'archivo-local': 'Abra el sitio desde http://localhost:8765 para iniciar sesión',
      'sin-libreria': 'No se cargó la librería de Microsoft (vendor/msal-browser.min.js)',
    }[conexion] || '';

    if (!conectado && ENVIO.modo === 'real') ENVIO.modo = 'prueba';
    const real = ENVIO.modo === 'real';
    const etiqueta = $('#etiqueta-modo');
    etiqueta.textContent = real ? 'Envío real' : 'Modo prueba';
    etiqueta.classList.toggle('real', real);

    const aviso = $('#envio-modo');
    aviso.className = 'alerta ' + (real ? 'error' : 'info');
    if (real) {
      aviso.replaceChildren(crear('strong', { texto: 'Envío real: ' }),
        `cada empleado marcado recibirá su boleta desde ${Microsoft.correo()}.`);
    } else if (conectado) {
      aviso.replaceChildren(crear('strong', { texto: 'Modo simulado: ' }),
        `la sesión está iniciada con ${Microsoft.correo()}, pero los correos a empleados se simulan. Use "Enviar prueba" para recibir una boleta real en su correo.`);
    } else {
      aviso.replaceChildren(crear('strong', { texto: 'Modo prueba: ' }),
        'no se envía ningún correo. El envío se simula para revisar el proceso. Para enviar correos, inicie sesión con Microsoft (botón arriba a la derecha).');
    }

    $('#envio-conexion').hidden = !conectado;
    $('#modo-cuenta').textContent = Microsoft.correo();
    document.querySelectorAll('input[name="modo-envio"]').forEach(r => {
      r.checked = r.value === ENVIO.modo;
      r.disabled = Boolean(estado.envio && estado.envio.enviando);
    });
    if (conectado && !$('#prueba-para').value) $('#prueba-para').value = Microsoft.correo();
    if (estado.envio && !$('#vista-envio').hidden) {
      mostrarEnvio();
      if (conectado && estado.envio.revisionLista.estado === 'sin-revisar') revisarYaEnviados();
    }
  }

  $('#btn-microsoft').addEventListener('click', async () => {
    const conexion = Microsoft.estado();
    if (conexion === 'sin-configurar') {
      alert('Todavía no está conectado Microsoft 365: falta el Client ID y el Tenant ID en js/config.js. Vea MICROSOFT365.md.');
      return;
    }
    if (conexion === 'archivo-local') {
      alert('Para iniciar sesión, abra el sitio desde http://localhost:8765 (no con doble clic en index.html).');
      return;
    }
    if (conexion === 'sin-libreria') {
      alert('No se cargó la librería de Microsoft (vendor/msal-browser.min.js).');
      return;
    }
    if (conexion === 'conectado') {
      if (estado.envio && estado.envio.enviando) return;
      if (confirm(`¿Cerrar la sesión de ${Microsoft.correo()} en este sitio?`)) {
        await Microsoft.cerrarSesion();
        mostrarConexion();
      }
      return;
    }
    try {
      await Microsoft.iniciarSesion();
    } catch (e) {
      console.error(e);
      const mensaje = Microsoft.mensajeDeSesion(e);
      if (mensaje) alert(mensaje);
    }
    mostrarConexion();
  });

  // Cambiar a envío real: lo que antes solo se simuló vuelve a quedar pendiente para enviarse de verdad.
  document.querySelectorAll('input[name="modo-envio"]').forEach(radio => {
    radio.addEventListener('change', () => {
      ENVIO.modo = radio.value;
      if (estado.envio) {
        estado.envio.items.forEach(i => {
          if (i.estado === 'enviado' && i.simulado && ENVIO.modo === 'real') i.estado = 'pendiente';
        });
        if (ENVIO.modo === 'real') estado.envio.terminado = false; // la simulación terminada da paso al envío real
      }
      mostrarConexion();
    });
  });

  // Envía la boleta que se está viendo a una sola dirección, con "[PRUEBA]" en el asunto.
  $('#btn-prueba').addEventListener('click', async () => {
    const para = $('#prueba-para').value.trim();
    const resultado = $('#prueba-resultado');
    if (!CORREO_VALIDO.test(para)) {
      resultado.className = 'resultado-prueba error';
      resultado.textContent = 'Escriba un correo válido';
      return;
    }
    const fila = estado.envio.items[estado.envio.actual].fila;
    const boton = $('#btn-prueba');
    boton.disabled = true;
    resultado.className = 'resultado-prueba';
    resultado.textContent = 'Enviando…';
    try {
      await Microsoft.enviarCorreo({ para, nombre: para, asunto: `[PRUEBA] ${Boleta.asunto(fila)}`, html: Boleta.html(fila) });
      resultado.className = 'resultado-prueba ok';
      resultado.textContent = `Enviada a ${para} a las ${hora(new Date())} (boleta de ${fila.nombre})`;
    } catch (e) {
      resultado.className = 'resultado-prueba error';
      resultado.textContent = e.message;
    } finally {
      boton.disabled = false;
    }
  });
  $('#prueba-para').addEventListener('input', () => { $('#prueba-resultado').textContent = ''; });

  // Si se intenta cerrar o recargar la página durante un envío real, el navegador pide confirmación.
  window.addEventListener('beforeunload', e => {
    if (estado.envio && estado.envio.enviando && ENVIO.modo === 'real') e.preventDefault();
  });

  Microsoft.iniciar()
    .catch(e => console.error('No se pudo preparar Microsoft 365', e))
    .finally(mostrarConexion);
})();
