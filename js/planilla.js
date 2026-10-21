// Lógica de la planilla quincenal: leer el Excel (o CSV), completar valores vacíos y calcular las columnas
// automáticas (total ingresos, cargas, deducciones y salario líquido).
// Funciona en el navegador (window.Planilla) y en Node (require), para poder probarla con archivos reales.
(function (root) {
  'use strict';

  const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
    'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  // Porcentajes por defecto de las cargas laborales del trabajador. El usuario los puede cambiar en "Tasas".
  const TASAS_POR_DEFECTO = { sem: 5.5, ivm: 4.33, banco: 1 };

  // Columnas que se leen del archivo, con los nombres exactos de la plantilla (al comparar se ignoran tildes,
  // mayúsculas y espacios de más: "Cédula" = "CEDULA"). El orden es el de la plantilla.
  const CAMPOS = [
    { key: 'cedula', encabezado: 'CEDULA', etiqueta: 'Cédula', tipo: 'cedula' },
    { key: 'nombre', encabezado: 'NOMBRE COMPLETO', etiqueta: 'Nombre completo', tipo: 'texto' },
    { key: 'correo', encabezado: 'CORREO', etiqueta: 'Correo', tipo: 'correo' },
    { key: 'bruto', encabezado: 'SALARIO BRUTO QUINCENAL', etiqueta: 'Salario bruto', tipo: 'monto' },
    { key: 'faltas', encabezado: 'FALTAS E INCAPACIDADES', etiqueta: 'Faltas e incapacidades', tipo: 'monto' },
    { key: 'horasExtra', encabezado: 'HORAS EXTRA O FERIADO', etiqueta: 'Horas extra o feriado', tipo: 'monto' },
    { key: 'vales', encabezado: 'VALES O PRESTAMOS', etiqueta: 'Vales o préstamos', tipo: 'monto' },
    { key: 'renta', encabezado: 'RENTA', etiqueta: 'Renta', tipo: 'monto' },
    { key: 'puesto', encabezado: 'PUESTO', etiqueta: 'Puesto', tipo: 'texto' },
    { key: 'empresa', encabezado: 'EMPRESA', etiqueta: 'Empresa', tipo: 'texto' },
    { key: 'fechaInicio', encabezado: 'FECHA DE INGRESO', etiqueta: 'Fecha de ingreso', tipo: 'fecha' },
  ];

  // Columnas que calcula la página (nunca se toman del archivo).
  const CALCULADAS = [
    { key: 'liquido', encabezado: 'SALARIO LIQUIDO', etiqueta: 'Salario líquido', tipo: 'monto' },
    { key: 'periodo', encabezado: 'PERIODO ACTUAL', etiqueta: 'Periodo actual', tipo: 'texto' },
    { key: 'totalIngresos', encabezado: 'TOTAL INGRESOS', etiqueta: 'Total ingresos', tipo: 'monto' },
    { key: 'sem', encabezado: 'ENFERMEDAD Y MATERNIDAD', etiqueta: 'Enfermedad y maternidad', tipo: 'monto' },
    { key: 'ivm', encabezado: 'INVALIDEZ', etiqueta: 'Invalidez, vejez y muerte', tipo: 'monto' },
    { key: 'banco', encabezado: 'BANCO', etiqueta: 'Banco Popular', tipo: 'monto' },
    { key: 'totalCargas', encabezado: 'TOTAL CARGAS LABORALES', etiqueta: 'Total cargas laborales', tipo: 'monto' },
    { key: 'totalDeducciones', encabezado: 'TOTAL DEDUCCIONES', etiqueta: 'Total deducciones', tipo: 'monto' },
  ];

  // Excel de salida y tabla del paso 4, en el orden de la plantilla (sin ID) con el periodo antes de las automáticas.
  const ORDEN_SALIDA = ['cedula', 'nombre', 'correo', 'liquido', 'bruto', 'faltas', 'horasExtra', 'vales', 'renta',
    'puesto', 'empresa', 'fechaInicio', 'periodo', 'totalIngresos', 'sem', 'ivm', 'banco', 'totalCargas', 'totalDeducciones'];
  const COLUMNAS_SALIDA = ORDEN_SALIDA.map(key => CAMPOS.concat(CALCULADAS).find(c => c.key === key));

  // Grupos de la fila superior, como en la plantilla: [título, primera columna, última columna].
  const GRUPOS = [
    ['IDENTIFICACIÓN', 'cedula', 'correo'],
    ['CONCEPTOS DE PAGO', 'liquido', 'horasExtra'],
    ['DEDUCCIONES', 'vales', 'fechaInicio'],
    ['AUTOMÁTICO', 'periodo', 'totalDeducciones'],
  ];

  // Encabezados de columnas calculadas (de esta plantilla y de las anteriores): se muestran como
  // "la página la calcula" y nunca se leen.
  const ENCABEZADOS_CALCULADOS = CALCULADAS.map(c => c.encabezado).concat(['TOTAL DEDDUCIONES']);

  // ---------- Utilidades ----------

  function normalizar(valor) {
    if (valor == null) return '';
    return String(valor).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
  }

  function estaVacio(v) {
    return v == null || (typeof v === 'string' && v.trim() === '');
  }

  function letraColumna(indice) {
    let n = indice + 1, s = '';
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // Filas de una hoja de SheetJS como matriz, siempre desde A1 para que las letras de columna coincidan.
  // Se lee el valor guardado de cada celda: fechas como número de serie (leerFecha las convierte), fórmulas con
  // su resultado y errores (#N/A, #REF!) como vacío. No se usa sheet_to_json porque, al leer con cellStyles,
  // convierte en "fecha inválida" el texto de una celda con formato de fecha (por ejemplo "16/09/2024").
  // Límite de filas y columnas por si la hoja tiene formato aplicado hasta el final de Excel.
  const MAXIMO_FILAS = 5000;
  const MAXIMO_COLUMNAS = 100;

  function filasDeHoja(XLSX, hoja) {
    if (!hoja || !hoja['!ref']) return [];
    const rango = XLSX.utils.decode_range(hoja['!ref']);
    const filas = [];
    for (let r = 0; r <= Math.min(rango.e.r, MAXIMO_FILAS - 1); r++) {
      const fila = [];
      for (let c = 0; c <= Math.min(rango.e.c, MAXIMO_COLUMNAS - 1); c++) {
        const celda = hoja[XLSX.utils.encode_cell({ r, c })];
        fila.push(!celda || celda.t === 'e' || celda.v === undefined ? null : celda.v);
      }
      filas.push(fila);
    }
    return filas;
  }

  // ---------- Lectura ----------

  // Los montos son colones enteros y positivos. Se aceptan escritos como texto, con símbolo y separadores:
  // "₡ 450,000", "? 450.000" (así sale el ₡ en un CSV), "450 000", "CRC 450000". Los separadores de miles se
  // reconocen por los grupos de 3 cifras; si hay punto y coma a la vez, el último es el decimal.
  // Los negativos ("-5" o "(5)") y los decimales se marcan para revisión.
  function leerMonto(v) {
    if (estaVacio(v)) return { valor: 0, relleno: true };
    if (typeof v === 'number') {
      if (!isFinite(v)) return { valor: 0, invalido: true };
      return { valor: v, convertido: false, negativo: v < 0, decimal: !Number.isInteger(v) };
    }
    const original = String(v).trim();
    const sinMoneda = original.replace(/^\s*(CRC|COL)\b/i, '');
    if (!/\d/.test(original)) {
      // Solo símbolos ("₡ -", "?") cuentan como celda vacía; texto como "pendiente" no es un número.
      return /[A-Za-zÀ-ÿ]/.test(sinMoneda) ? { valor: 0, invalido: true } : { valor: 0, relleno: true };
    }
    if (/[A-Za-zÀ-ÿ]/.test(sinMoneda)) return { valor: 0, invalido: true };

    const negativo = original.includes('-') || /^\(.*\)$/.test(original);
    const texto = original.replace(/[^\d.,]/g, '');
    const punto = texto.lastIndexOf('.');
    const coma = texto.lastIndexOf(',');
    let valor;
    if (punto >= 0 && coma >= 0) {
      const decimal = punto > coma ? '.' : ',';
      valor = Number(texto.split(decimal === '.' ? ',' : '.').join('').replace(decimal, '.'));
    } else if (/^\d{1,3}([.,]\d{3})+$/.test(texto)) {
      valor = Number(texto.replace(/[.,]/g, ''));
    } else {
      valor = Number(texto.replace(',', '.'));
    }
    if (isNaN(valor)) return { valor: 0, invalido: true };
    if (negativo) valor = -valor;
    return { valor, convertido: !/^-?\d+$/.test(original), negativo, decimal: !Number.isInteger(valor) };
  }

  // Cédula nacional: 9 dígitos → 0-0000-0000 (también corrige "01-1802-0952" y el formato viejo "1-802-952").
  // Residentes (DIMEX): 11 o 12 dígitos, se dejan como se escribieron (con o sin guiones).
  // Cualquier otro largo se deja como viene y se marca.
  function leerCedula(v) {
    if (estaVacio(v)) return { texto: '' };
    const original = typeof v === 'number' ? String(Math.round(v)) : String(v).trim();
    let digitos = original.replace(/\D/g, '');
    const partes = original.match(/^0?(\d)-(\d{1,4})-(\d{1,4})$/);
    if (partes) digitos = partes[1] + partes[2].padStart(4, '0') + partes[3].padStart(4, '0');
    if (digitos.length === 10 && digitos[0] === '0') digitos = digitos.slice(1);
    if (digitos.length === 9) return { texto: `${digitos[0]}-${digitos.slice(1, 5)}-${digitos.slice(5)}` };
    if (digitos.length === 11 || digitos.length === 12) return { texto: /^[\d-]+$/.test(original) ? original : digitos };
    return { texto: original, invalida: `La cédula "${original}" tiene ${digitos.length} dígitos (nacional: 9; residente/DIMEX: 11 o 12)` };
  }

  function formatearCedula(v) {
    return leerCedula(v).texto;
  }

  // Fechas: de Excel (número de serie) o escritas como texto: "16/09/2024", "16-09-24", "2024-09-16",
  // "01-dic.-17", "1 de diciembre de 2017". Día primero (como en Costa Rica); una fecha imposible
  // (por ejemplo "11/16/2024", con el mes primero) no se adivina y queda vacía.
  const MESES_CORTOS = { ENE: 1, JAN: 1, FEB: 2, MAR: 3, ABR: 4, APR: 4, MAY: 5, JUN: 6, JUL: 7, AGO: 8, AUG: 8,
    SET: 9, SEP: 9, OCT: 10, NOV: 11, DIC: 12, DEC: 12 };

  function crearFecha(anio, mes, dia) {
    if (anio < 100) anio += anio < 50 ? 2000 : 1900;
    const d = new Date(Date.UTC(anio, mes - 1, dia));
    return d.getUTCFullYear() === anio && d.getUTCMonth() === mes - 1 && d.getUTCDate() === dia ? d : null;
  }

  function leerFecha(v) {
    if (estaVacio(v)) return null;
    if (v instanceof Date) return isNaN(v) ? null : v;
    if (typeof v === 'number') return v > 0 ? new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000) : null;
    const texto = normalizar(v);
    let m = texto.match(/^(\d{4}) (\d{1,2}) (\d{1,2})$/);
    if (m) return crearFecha(+m[1], +m[2], +m[3]);
    m = texto.match(/^(\d{1,2}) (\d{1,2}) (\d{2}|\d{4})$/);
    if (m) return crearFecha(+m[3], +m[2], +m[1]);
    m = texto.match(/^(\d{1,2}) (?:DE )?([A-Z]{3})[A-Z]* (?:DE )?(\d{2}|\d{4})$/);
    if (m && MESES_CORTOS[m[2]]) return crearFecha(+m[3], MESES_CORTOS[m[2]], +m[1]);
    return null;
  }

  // Columna de cada campo en una fila de encabezados, comparando con el nombre exacto de la plantilla.
  function mapearColumnas(fila) {
    const encabezados = Array.from(fila, normalizar);
    const mapa = {};
    CAMPOS.forEach(campo => {
      const c = encabezados.indexOf(normalizar(campo.encabezado));
      if (c >= 0) mapa[campo.key] = c;
    });
    return mapa;
  }

  // La fila de encabezados puede estar en cualquiera de las primeras 50 filas (la plantilla tiene títulos
  // encima). Es la fila que tiene CEDULA y NOMBRE COMPLETO; si hay varias, la que tiene más columnas.
  const FILAS_A_REVISAR = 50;

  function buscarEncabezado(filas) {
    let mejor = null;
    for (let i = 0; i < Math.min(filas.length, FILAS_A_REVISAR); i++) {
      const mapa = mapearColumnas(filas[i] || []);
      if (mapa.cedula == null || mapa.nombre == null) continue;
      if (!mejor || Object.keys(mapa).length > Object.keys(mejor.mapa).length) mejor = { fila: i, mapa };
    }
    return mejor;
  }

  function filaMasParecida(filas) {
    let mejor = null;
    for (let i = 0; i < Math.min(filas.length, FILAS_A_REVISAR); i++) {
      const mapa = mapearColumnas(filas[i] || []);
      if (Object.keys(mapa).length && (!mejor || Object.keys(mapa).length > Object.keys(mejor.mapa).length)) mejor = { fila: i, mapa };
    }
    return mejor;
  }

  // Siempre se lee la primera hoja del archivo y debe tener todas las columnas de la plantilla. Si no, se
  // devuelve un error que dice exactamente qué falta.
  function primeraHoja(XLSX, libro) {
    const nombre = libro.SheetNames[0];
    const hoja = libro.Sheets[nombre];
    const filas = filasDeHoja(XLSX, hoja);
    const encabezado = buscarEncabezado(filas);
    // Si no hay una fila con CEDULA y NOMBRE COMPLETO, se usa la que más se parece para decir qué columnas faltan.
    const comparada = encabezado || filaMasParecida(filas);
    if (!comparada) {
      return { error: `La primera hoja ("${nombre}") no tiene ninguno de los encabezados de la plantilla.` };
    }
    const faltantes = CAMPOS.filter(c => comparada.mapa[c.key] == null);
    if (faltantes.length) {
      return { error: `A la hoja "${nombre}" le ${faltantes.length === 1 ? 'falta la columna' : 'faltan las columnas'} ` +
        `${faltantes.map(c => c.encabezado).join(', ')}. Los nombres deben ser los de la plantilla.` };
    }
    // Filas ocultas en Excel (también las que esconde un filtro). Requiere leer el libro con { cellStyles: true }.
    const ocultas = new Set();
    (hoja['!rows'] || []).forEach((fila, i) => { if (fila && fila.hidden) ocultas.add(i); });
    return { hoja: { nombre, filas, encabezado, ocultas, esCsv: Boolean(libro.esCsv) } };
  }

  const CORREO_VALIDO = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  // Datos que pueden faltar sin detener el proceso: se completan con "N/A" y se avisan en amarillo.
  const OPCIONALES = new Set(['fechaInicio', 'puesto', 'empresa']);
  const NO_DISPONIBLE = 'N/A';
  const FILA_DE_TOTALES = /^(SUB ?)?TOTAL(ES)?( |$)/;

  // Tipos de nota: "relleno"/"invalido"/"ajuste"/"na" = se ajustó el valor (amarillo); "vacio"/"error" = problema (rojo).
  function esProblema(nota) {
    return nota.tipo === 'vacio' || nota.tipo === 'error';
  }

  // Filas que no son empleados: totales al final de la planilla, o notas sueltas sin cédula, correo ni montos.
  // "TOTAL" se busca en cédula, nombre y columnas no reconocidas (no en puesto o empresa, que podrían contenerlo).
  function motivoParaIgnorar(fila, obtener, mapa) {
    // Filas de prueba que la plantilla trae marcadas, como "AA PRUEBA NO TOCAR".
    if (/\bNO (TOCAR|USAR)\b/.test(normalizar(obtener('nombre')))) return 'Fila de prueba ("NO TOCAR")';
    const otrasColumnas = new Set(['correo', 'puesto', 'empresa'].map(k => mapa[k]).filter(c => c != null));
    const dice = (v, c) => !otrasColumnas.has(c) && typeof v === 'string' && FILA_DE_TOTALES.test(normalizar(v));
    if (fila.some(dice)) return 'Fila de totales';
    const sinMontos = CAMPOS.filter(c => c.tipo === 'monto').every(c => leerMonto(obtener(c.key)).valor === 0);
    if (estaVacio(obtener('cedula')) && estaVacio(obtener('correo')) && sinMontos) {
      return 'Sin cédula, correo ni montos (parece una nota)';
    }
    return null;
  }

  function textoDeFila(fila) {
    const texto = fila.filter(v => !estaVacio(v)).map(v => String(v).trim()).join(' · ');
    return texto.length > 80 ? texto.slice(0, 79) + '…' : texto;
  }

  // Cédulas y correos que aparecen en más de una fila: se marcan en todas, indicando dónde más están.
  function marcarRepetidos(empleados, key, texto) {
    const grupos = new Map();
    empleados.forEach(e => {
      if (!e[key]) return;
      const clave = key === 'cedula' ? (e.cedula.replace(/\D/g, '') || e.cedula) : e[key];
      if (!grupos.has(clave)) grupos.set(clave, []);
      grupos.get(clave).push(e);
    });
    grupos.forEach(lista => {
      if (lista.length < 2) return;
      lista.forEach(e => {
        const otras = lista.filter(o => o !== e).map(o => o.filaExcel).join(', ');
        e.notas.push({ campo: key, tipo: 'error', codigo: 'repetido', texto: `${texto} (también en la fila ${otras})` });
      });
    });
  }

  function leerPlanilla(hoja) {
    const { filas, encabezado } = hoja;
    const mapa = encabezado.mapa;
    const empleados = [];
    const ignoradas = [];

    for (let i = encabezado.fila + 1; i < filas.length; i++) {
      const fila = filas[i] || [];
      const obtener = key => (mapa[key] == null ? undefined : fila[mapa[key]]);
      // Filas sin cédula, nombre ni correo (por ejemplo, filas de la plantilla con fórmulas en 0) se saltan sin aviso.
      if (['cedula', 'nombre', 'correo'].every(k => estaVacio(obtener(k)))) continue;
      if (normalizar(obtener('cedula')) === 'CEDULA') continue;
      const motivo = hoja.ocultas && hoja.ocultas.has(i) ? 'Fila oculta en el Excel' : motivoParaIgnorar(fila, obtener, mapa);
      if (motivo) {
        ignoradas.push({ filaExcel: i + 1, texto: textoDeFila(fila), motivo });
        continue;
      }

      const empleado = { filaExcel: i + 1, notas: [] };
      const anotar = (campo, tipo, texto, codigo) => empleado.notas.push({ campo: campo.key, tipo, texto, codigo });

      CAMPOS.forEach(campo => {
        const valor = obtener(campo.key);
        if (campo.tipo === 'monto') {
          const monto = leerMonto(valor);
          empleado[campo.key] = monto.valor;
          if (monto.invalido) anotar(campo, 'invalido', `"${valor}" no es un número; se usó 0`);
          else if (monto.relleno) anotar(campo, 'relleno', 'Celda vacía; se completó con 0');
          else {
            // En un CSV todo viene como texto: convertirlo es lo normal y no se marca.
            if (monto.convertido && !hoja.esCsv && !monto.negativo && !monto.decimal) anotar(campo, 'ajuste', `Venía como texto "${valor}"; se leyó como ${monto.valor}`);
            if (monto.negativo) anotar(campo, 'error', `Monto negativo (${monto.valor}); los montos deben ser positivos`, 'negativo');
            if (monto.decimal) anotar(campo, 'error', `Tiene decimales (${monto.valor}); los montos deben ser colones enteros`, 'decimal');
          }
        } else if (campo.tipo === 'fecha') {
          // Fecha vacía o que no se entiende → "N/A" con aviso amarillo (no impide continuar).
          const fecha = leerFecha(valor);
          empleado[campo.key] = fecha || NO_DISPONIBLE;
          if (!fecha && !estaVacio(valor)) anotar(campo, 'na', `"${valor}" no es una fecha válida; se puso N/A`, 'na');
          else if (!fecha) anotar(campo, 'na', 'Sin fecha de inicio; se puso N/A', 'na');
        } else {
          let texto = estaVacio(valor) ? '' : String(valor).trim().replace(/\s+/g, ' ');
          if (campo.tipo === 'cedula') {
            const cedula = leerCedula(valor);
            texto = cedula.texto;
            if (cedula.invalida) anotar(campo, 'error', cedula.invalida, 'cedula');
          }
          if (campo.tipo === 'correo') texto = texto.toLowerCase();
          // Puesto y empresa vacíos → "N/A" con aviso amarillo. Cédula, nombre y correo vacíos → rojo.
          if (!texto && OPCIONALES.has(campo.key)) {
            texto = NO_DISPONIBLE;
            anotar(campo, 'na', 'Celda vacía; se puso N/A', 'na');
          }
          empleado[campo.key] = texto;
          if (!texto) anotar(campo, 'vacio', 'Celda vacía');
          else if (campo.tipo === 'correo' && texto && !CORREO_VALIDO.test(texto)) {
            anotar(campo, 'error', `"${texto}" no es un correo válido`, 'formato');
          }
        }
      });
      empleados.push(empleado);
    }

    marcarRepetidos(empleados, 'cedula', 'Cédula repetida');
    marcarRepetidos(empleados, 'correo', 'Correo repetido');

    // Columnas con encabezado que no se usan (ID, nacionalidad, antigüedad…). Las de la plantilla que la
    // página calcula por su cuenta (PERIODO ACTUAL … SALARIO LIQUIDO) se marcan como recalculadas.
    const usadas = new Set(Object.values(mapa));
    const calculadas = new Set(ENCABEZADOS_CALCULADOS.map(normalizar));
    const columnasIgnoradas = (filas[encabezado.fila] || [])
      .map((valor, c) => ({ c, texto: estaVacio(valor) ? '' : String(valor).trim() }))
      .filter(x => x.texto && !usadas.has(x.c))
      .map(x => ({ columna: letraColumna(x.c), encabezado: x.texto, recalculada: calculadas.has(normalizar(x.texto)) }));

    return {
      ignoradas,
      columnasIgnoradas,
      hoja: hoja.nombre,
      filaEncabezado: encabezado.fila + 1,
      columnas: CAMPOS.map(campo => ({
        campo,
        columna: mapa[campo.key] == null ? null : letraColumna(mapa[campo.key]),
      })),
      empleados,
    };
  }

  // ---------- Quincenas ----------

  function moverQuincena(p, pasos) {
    const indice = p.anio * 24 + p.mes * 2 + (p.q - 1) + pasos;
    return { anio: Math.floor(indice / 24), mes: Math.floor((indice % 24) / 2), q: (indice % 2) + 1 };
  }

  // Días 1–15 = primera quincena; 16–fin de mes = segunda.
  function opcionesQuincena(hoy) {
    const actual = { anio: hoy.getFullYear(), mes: hoy.getMonth(), q: hoy.getDate() <= 15 ? 1 : 2 };
    return [
      Object.assign({ tipo: 'Actual' }, actual),
      Object.assign({ tipo: 'Siguiente' }, moverQuincena(actual, 1)),
      Object.assign({ tipo: 'Anterior' }, moverQuincena(actual, -1)),
    ];
  }

  function textoQuincena(p) {
    return `${p.q === 1 ? 'PRIMERA' : 'SEGUNDA'} QUINCENA DE ${MESES[p.mes]} DE ${p.anio}`;
  }

  // Nombre del mes con mayúscula inicial ("Octubre"), para las pantallas y el nombre del archivo.
  function nombreMes(mes) {
    return MESES[mes].charAt(0) + MESES[mes].slice(1).toLowerCase();
  }

  // "Segunda quincena de Octubre de 2026". En el Excel se usa textoQuincena, en mayúsculas como la plantilla.
  function etiquetaQuincena(p) {
    return `${p.q === 1 ? 'Primera' : 'Segunda'} quincena de ${nombreMes(p.mes)} de ${p.anio}`;
  }

  // ---------- Cálculos ----------

  // Total ingresos = bruto − faltas e incapacidades + horas extra. Las cargas se calculan sobre el total de
  // ingresos y cada una se redondea a colones enteros antes de sumar, para que la boleta cuadre línea por línea.
  function calcularEmpleado(e, tasas) {
    const totalIngresos = e.bruto - e.faltas + e.horasExtra;
    const sem = Math.round(totalIngresos * tasas.sem / 100);
    const ivm = Math.round(totalIngresos * tasas.ivm / 100);
    const banco = Math.round(totalIngresos * tasas.banco / 100);
    const totalCargas = sem + ivm + banco;
    const totalDeducciones = totalCargas + e.vales + e.renta;
    return { totalIngresos, sem, ivm, banco, totalCargas, totalDeducciones, liquido: totalIngresos - totalDeducciones };
  }

  function calcularPlanilla(lectura, periodo, tasas) {
    return lectura.empleados.map(e => Object.assign({}, e, calcularEmpleado(e, tasas), {
      periodo: textoQuincena(periodo),
    }));
  }

  const CLAVES_TOTALES = ['bruto', 'totalIngresos', 'totalCargas', 'totalDeducciones', 'liquido'];

  function totales(filas) {
    const nuevo = nombre => CLAVES_TOTALES.reduce((t, k) => (t[k] = 0, t), { empresa: nombre, empleados: 0 });
    const general = nuevo('Total');
    const porEmpresa = new Map();
    filas.forEach(f => {
      const nombre = f.empresa || '(sin empresa)';
      if (!porEmpresa.has(nombre)) porEmpresa.set(nombre, nuevo(nombre));
      [porEmpresa.get(nombre), general].forEach(t => {
        t.empleados++;
        CLAVES_TOTALES.forEach(k => { t[k] += f[k]; });
      });
    });
    return { porEmpresa: Array.from(porEmpresa.values()), general };
  }

  // ---------- Excel de salida ----------

  // Formato personalizado: la celda guarda solo el número (201052) y Excel lo muestra como ₡201,052.
  const FORMATO_MONTO = '"₡"#,##0;-"₡"#,##0';
  const AZUL_OSCURO = 'FF1F3A5F';
  const AMARILLO = 'FFFFF4CC';
  const ROJO_SUAVE = 'FFFDECEA';

  function anchoColumna(col) {
    if (col.key === 'nombre') return 34;
    if (col.key === 'correo') return 28;
    if (col.key === 'periodo') return 40;
    if (col.tipo === 'monto') return Math.max(14, col.encabezado.length * 0.9);
    return 16;
  }

  // Una sola hoja, con la tabla "SalarioV1" (para Power Automate) y las columnas en el orden de la plantilla.
  // Las celdas ajustadas van en amarillo y los problemas en rojo; la explicación está en la página.
  // No se agregan notas a las celdas: ExcelJS escribe las notas junto con la tabla en un orden que Excel
  // rechaza ("estamos intentando recuperar…") y la hoja queda vacía al abrirla.
  function construirExcel(ExcelJS, datos) {
    const { filas } = datos;
    const libro = new ExcelJS.Workbook();
    libro.creator = 'Boletas de pago';
    libro.created = new Date();

    const hoja = libro.addWorksheet('Zodiac Salarios', { views: [{ state: 'frozen', xSplit: 2, ySplit: 2 }] });
    const letra = key => letraColumna(COLUMNAS_SALIDA.findIndex(c => c.key === key));
    GRUPOS.forEach(([texto, desde, hasta]) => {
      hoja.mergeCells(`${letra(desde)}1:${letra(hasta)}1`);
      const celda = hoja.getCell(`${letra(desde)}1`);
      celda.value = texto;
      celda.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_OSCURO } };
      celda.alignment = { horizontal: 'center' };
    });

    hoja.addTable({
      name: 'SalarioV1',
      ref: 'A2',
      headerRow: true,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: COLUMNAS_SALIDA.map(c => ({ name: c.encabezado, filterButton: true })),
      rows: filas.map(f => COLUMNAS_SALIDA.map(c => (f[c.key] == null ? '' : f[c.key]))),
    });

    COLUMNAS_SALIDA.forEach((c, i) => {
      const columna = hoja.getColumn(i + 1);
      columna.width = anchoColumna(c);
      if (c.tipo === 'monto') columna.numFmt = FORMATO_MONTO;
      if (c.tipo === 'fecha') columna.numFmt = 'dd/mm/yyyy';
    });
    hoja.getRow(2).alignment = { wrapText: true, vertical: 'middle' };
    hoja.getRow(2).height = 32;

    const pintar = (r, key, color) => {
      hoja.getCell(r + 3, COLUMNAS_SALIDA.findIndex(c => c.key === key) + 1)
        .fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
    };
    filas.forEach((f, r) => {
      new Set(f.notas.map(n => n.campo)).forEach(campo => {
        pintar(r, campo, f.notas.some(n => n.campo === campo && esProblema(n)) ? ROJO_SUAVE : AMARILLO);
      });
      if (f.liquido <= 0) pintar(r, 'liquido', ROJO_SUAVE);
    });
    return libro;
  }

  // ---------- Plantilla para descargar ----------

  // Archivo vacío con el formato que la página entiende: solo las columnas que se llenan (lo demás lo calcula
  // la página), una fila de ejemplo que se ignora sola ("NO TOCAR"), formatos de monto/fecha/cédula y
  // validaciones para que Excel avise si se escribe un monto negativo o con decimales, o una fecha inválida.
  const FILAS_PLANTILLA = 500;
  const GRUPOS_PLANTILLA = [
    ['IDENTIFICACIÓN', 'cedula', 'correo'],
    ['CONCEPTOS DE PAGO', 'bruto', 'horasExtra'],
    ['DEDUCCIONES', 'vales', 'renta'],
    ['PUESTO Y EMPRESA', 'puesto', 'fechaInicio'],
  ];

  function construirPlantilla(ExcelJS) {
    const libro = new ExcelJS.Workbook();
    libro.creator = 'Boletas de pago';
    const hoja = libro.addWorksheet('Planilla', { views: [{ state: 'frozen', xSplit: 2, ySplit: 2 }] });
    const indice = key => CAMPOS.findIndex(c => c.key === key) + 1;
    const letra = key => letraColumna(indice(key) - 1);
    const ultima = 2 + FILAS_PLANTILLA;

    GRUPOS_PLANTILLA.forEach(([texto, desde, hasta]) => {
      hoja.mergeCells(`${letra(desde)}1:${letra(hasta)}1`);
      const celda = hoja.getCell(`${letra(desde)}1`);
      celda.value = texto;
      celda.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL_OSCURO } };
      celda.alignment = { horizontal: 'center' };
    });
    const encabezado = hoja.getRow(2);
    CAMPOS.forEach((c, i) => {
      const celda = encabezado.getCell(i + 1);
      celda.value = c.encabezado;
      celda.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E5C94' } };
      celda.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    encabezado.height = 32;

    // Fila de ejemplo: se ve cómo llenar cada columna y la página la ignora por el "NO TOCAR".
    const ejemplo = { cedula: '1-0000-0000', nombre: 'AA EJEMPLO NO TOCAR (se ignora)', correo: 'ejemplo@zodiac.co.cr',
      bruto: 300000, faltas: 5000, horasExtra: 15000, vales: 20000, renta: 2500, puesto: 'Bodeguero',
      empresa: 'CORP.BEXIM', fechaInicio: new Date(Date.UTC(2024, 0, 15)) };
    const filaEjemplo = hoja.getRow(3);
    CAMPOS.forEach((c, i) => {
      const celda = filaEjemplo.getCell(i + 1);
      celda.value = ejemplo[c.key];
      celda.font = { italic: true, color: { argb: 'FF8A94A3' } };
    });

    CAMPOS.forEach((c, i) => {
      const columna = hoja.getColumn(i + 1);
      columna.width = c.key === 'nombre' ? 34 : c.key === 'correo' ? 30 : c.tipo === 'monto' ? 16 : 16;
      for (let r = 3; r <= ultima; r++) {
        const celda = hoja.getCell(r, i + 1);
        if (c.tipo === 'monto') celda.numFmt = FORMATO_MONTO;
        if (c.tipo === 'fecha') celda.numFmt = 'dd/mm/yyyy';
        if (c.tipo === 'cedula') celda.numFmt = '@'; // texto: no se pierden ceros ni guiones
      }
    });
    hoja.dataValidations.add(`${letra('bruto')}3:${letra('renta')}${ultima}`, {
      type: 'whole', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [0], showErrorMessage: true,
      errorTitle: 'Monto no válido', error: 'Escriba el monto en colones, como número entero y positivo (sin decimales ni signo menos).',
    });
    hoja.dataValidations.add(`${letra('fechaInicio')}3:${letra('fechaInicio')}${ultima}`, {
      type: 'date', operator: 'between', allowBlank: true,
      formulae: [new Date(Date.UTC(1950, 0, 1)), new Date(Date.UTC(2100, 11, 31))], showErrorMessage: true,
      errorTitle: 'Fecha no válida', error: 'Escriba la fecha como día/mes/año, por ejemplo 15/01/2024.',
    });

    const ayuda = libro.addWorksheet('Instrucciones');
    ayuda.getColumn(1).width = 30;
    ayuda.getColumn(2).width = 90;
    const lineas = [
      ['Cómo llenar la planilla', ''],
      ['', ''],
      ['Una fila por empleado', 'Empiece en la fila 4 de la hoja "Planilla". La fila 3 es un ejemplo y la página la ignora.'],
      ['Encabezados', 'No cambie los nombres de las columnas. Puede agregar columnas propias: la página las ignora.'],
      ['Cédula', 'Nacional: 1-2345-6789. Residentes (DIMEX): los 11 o 12 dígitos, con o sin guiones.'],
      ['Correo', 'El correo donde el empleado recibe su boleta.'],
      ['Montos', 'En colones, números enteros y positivos, sin ₡ ni decimales. Si no aplica, escriba 0 (una celda vacía también cuenta como 0, pero la página la marca en amarillo para revisar).'],
      ['Faltas e incapacidades', 'En positivo: la página lo resta del salario y en la boleta se muestra en negativo.'],
      ['Horas extra o feriado', 'Se suma al salario bruto.'],
      ['Vales o préstamos y Renta', 'Se restan como deducciones.'],
      ['Fecha de ingreso', 'Día/mes/año, por ejemplo 15/01/2024.'],
      ['Lo que calcula la página', 'Total ingresos, Enfermedad y maternidad, Invalidez, Banco Popular, Total cargas, Total deducciones y Salario líquido. No hace falta llenarlos.'],
      ['Quincena', 'Se elige en la página (paso 3); no hace falta escribirla en el archivo.'],
    ];
    lineas.forEach(([a, b], i) => {
      const fila = ayuda.addRow([a, b]);
      if (i === 0) fila.font = { bold: true, size: 14, color: { argb: 'FF1F3A5F' } };
      else if (a) fila.getCell(1).font = { bold: true };
      fila.getCell(2).alignment = { wrapText: true, vertical: 'top' };
    });
    return libro;
  }

  const api = {
    MESES, TASAS_POR_DEFECTO, CAMPOS, CALCULADAS, COLUMNAS_SALIDA, GRUPOS,
    normalizar, formatearCedula, leerMonto, leerFecha, esProblema,
    primeraHoja, leerPlanilla,
    opcionesQuincena, textoQuincena, etiquetaQuincena, nombreMes,
    calcularEmpleado, calcularPlanilla, totales, construirExcel, construirPlantilla,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Planilla = api;
})(this);
