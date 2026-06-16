/**
 * GEMELO DIGITAL DE MERCADO — Backend (Google Apps Script)  ·  v1
 * --------------------------------------------------------------------------
 * Paper trading "autónomo" con datos reales (Yahoo Finance v8).
 * Scoring transparente SIN IA (0 tokens) + overlay geopolítico OPCIONAL con IA.
 * Trazabilidad completa en Google Sheets. Impuestos OFF (ganancias = líquido neto).
 * Híbrido: te propone operaciones y tú apruebas; AUTO_APROBAR=true => autónomo total.
 *
 * INSTALACIÓN (ver pasos en el chat):
 *  1) Crea una hoja de cálculo nueva en Google Sheets.
 *  2) Extensiones > Apps Script. Pega TODO este archivo. Guarda.
 *  3) Ejecuta la función  setup()  una vez (autoriza permisos).
 *  4) Recarga la hoja: aparece el menú "Gemelo 📈".
 * --------------------------------------------------------------------------
 */

// ===================== CONFIG POR DEFECTO (editable en pestaña Config) =====================
var DEFAULTS = {
  LIQUIDO_INICIAL: 2000,      // € de partida en el gemelo
  COMISION: 1,                // € por orden (Trade Republic). Pon 0 si no quieres costes
  MAX_POSICIONES: 6,          // nº máximo de posiciones simultáneas
  CAP_POSICION_PCT: 25,       // % máx del patrimonio del gemelo por posición
  BUFFER_LIQUIDEZ_PCT: 8,     // % de caja que se intenta mantener sin invertir
  UMBRAL_COMPRA: 60,          // score mínimo para abrir compra
  UMBRAL_VENTA: 45,           // score por debajo del cual se cierra
  STOP_LOSS_PCT: -12,         // venta forzada por pérdida desde entrada
  TAKE_PROFIT_PCT: 25,        // venta forzada por objetivo de ganancia
  // --- Gestión de riesgo (B) ---
  FILTRO_REGIMEN: true,       // solo compra si el benchmark está por encima de su MA200 (mercado alcista)
  MAX_POR_SECTOR: 2,          // máximo de posiciones por sector (diversificación)
  MAX_DRAWDOWN_PCT: -20,      // si el patrimonio cae más que esto desde su pico, deja de comprar (defensa)
  AUTO_APROBAR: false,        // false = híbrido (tú apruebas) · true = autónomo total
  USAR_IA: false,             // overlay geopolítico con IA (consume tokens). Apagado por defecto
  ANTHROPIC_API_KEY: '',      // pega tu clave SOLO si activas USAR_IA
  // --- Backtest / benchmark ---
  BENCHMARK: '^GSPC',         // índice de referencia (S&P 500). Alternativas: 'SPY', tu ETF 'IUSQ.DE'
  BACKTEST_ANOS: 2,           // años de histórico a simular
  BACKTEST_REBALANCEO_DIAS: 5,// cada cuántos días de bolsa revisa la cartera (5 ≈ semanal)
  // Pesos del score (se normalizan; ajústalos a tu gusto)
  W_MOMENTUM: 45,
  W_TENDENCIA: 25,
  W_POS52W: 15,
  W_SENTIMIENTO: 15,
  W_FUERZA: 20             // fuerza relativa frente al benchmark (¿bate al mercado?)
};

var TABS = {
  CONFIG:'Config', UNIVERSO:'Universo', DATOS:'Datos', SENT:'SentimientoSector',
  LIQ:'GemeloLiquidez', POS:'GemeloPosiciones', OPS:'Operaciones',
  HIST:'Historico', REAL:'CarteraReal', RECO:'Recomendaciones'
};

// Universo inicial relevante en Trade Republic (Ticker Yahoo, Nombre, Sector). Edítalo libremente.
var UNIVERSO_SEED = [
  ['AAPL','Apple','Tecnología'], ['MSFT','Microsoft','Tecnología'],
  ['NVDA','NVIDIA','Tecnología'], ['GOOGL','Alphabet','Tecnología'],
  ['AMZN','Amazon','Consumo'],    ['META','Meta','Tecnología'],
  ['TSLA','Tesla','Automoción'],  ['INTC','Intel','Tecnología'],
  ['AMD','AMD','Tecnología'],     ['ASML.AS','ASML','Tecnología'],
  ['SAP.DE','SAP','Tecnología'],  ['SIE.DE','Siemens','Industrial'],
  ['ALV.DE','Allianz','Financiero'], ['AIR.PA','Airbus','Industrial'],
  ['MC.PA','LVMH','Consumo'],     ['NESN.SW','Nestlé','Consumo'],
  ['SAN.MC','Santander','Financiero'], ['BBVA.MC','BBVA','Financiero'],
  ['ITX.MC','Inditex','Consumo'], ['IBE.MC','Iberdrola','Energía'],
  ['REP.MC','Repsol','Energía'],  ['TEF.MC','Telefónica','Telecom'],
  ['LMT','Lockheed Martin','Defensa'], ['RTX','RTX','Defensa'],
  ['XOM','ExxonMobil','Energía'], ['CVX','Chevron','Energía']
];

var SENT_SEED = ['Tecnología','Consumo','Automoción','Industrial','Financiero',
                 'Energía','Telecom','Defensa','Salud','Materiales'];

// Tus posiciones reales (prerellenadas: rellena Unidades/PrecioMedio/Cuota). + filas vacías para añadir.
var REAL_SEED = [
  ['IUSQ.DE','iShares Core MSCI World (revisar)','ETF','', '', ''],
  ['ELG.DE','ETF (revisar nombre/ISIN)','ETF','', '', ''],
  ['','','', '', '', ''],
  ['','','', '', '', ''],
  ['','','', '', '', ''],
  ['','','', '', '', '']
];

// ===================== MENÚ =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gemelo 📈')
    .addItem('① Actualizar datos (mercado)', 'actualizarDatos')
    .addItem('② Generar decisiones', 'generarDecisiones')
    .addItem('③ Ejecutar aprobadas', 'ejecutarAprobadas')
    .addItem('④ Evaluar cartera real + DCA', 'evaluarCarteraReal')
    .addItem('⑤ Snapshot patrimonio', 'tomarSnapshot')
    .addItem('📊 Backtest histórico', 'backtest')
    .addSeparator()
    .addItem('🌍 Actualizar sentimiento (IA)', 'actualizarSentimientoIA')
    .addItem('🎨 Aplicar estilos / selectores', 'aplicarEstilos')
    .addItem('⚙️ Crear / reparar hoja (setup)', 'setup')
    .addToUi();
}

// ===================== SETUP =====================
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // CONFIG
  var cfg = getOrCreateSheet(TABS.CONFIG);
  cfg.clearContents();
  var cfgRows = [['Clave','Valor']];
  for (var k in DEFAULTS) cfgRows.push([k, DEFAULTS[k]]);
  cfg.getRange(1,1,cfgRows.length,2).setValues(cfgRows);
  cfg.getRange(1,1,1,2).setFontWeight('bold'); cfg.setColumnWidth(1,220);

  // UNIVERSO
  var uni = getOrCreateSheet(TABS.UNIVERSO);
  uni.clearContents();
  uni.getRange(1,1,1,3).setValues([['Ticker','Nombre','Sector']]).setFontWeight('bold');
  uni.getRange(2,1,UNIVERSO_SEED.length,3).setValues(UNIVERSO_SEED);
  uni.setFrozenRows(1);

  // SENTIMIENTO SECTOR
  var sen = getOrCreateSheet(TABS.SENT);
  sen.clearContents();
  sen.getRange(1,1,1,4).setValues([['Sector','Sentimiento(-2..+2)','Nota','Actualizado']]).setFontWeight('bold');
  var senRows = SENT_SEED.map(function(s){ return [s, 0, 'neutral', '']; });
  sen.getRange(2,1,senRows.length,4).setValues(senRows);
  sen.setFrozenRows(1); sen.setColumnWidth(3,320);

  // DATOS (estructura; el formato heatmap se aplica en aplicarEstilos)
  var dat = getOrCreateSheet(TABS.DATOS);
  dat.clear();
  dat.getRange(1,1,1,17).setValues([[
    'Ticker','Nombre','Sector','Precio','Cambio%','MA50','MA200',
    'Mom1m%','Mom3m%','Mom6m%','Vol%','Pos52s','Sentim.','Score','Señal','Actualizado','FuerzaRel%'
  ]]).setFontWeight('bold');
  dat.setFrozenRows(1);

  // GEMELO LIQUIDEZ
  var liq = getOrCreateSheet(TABS.LIQ);
  liq.clearContents();
  liq.getRange(1,1,5,2).setValues([
    ['Concepto','Valor (€)'],
    ['Líquido disponible', DEFAULTS.LIQUIDO_INICIAL],
    ['P&L realizado acumulado', 0],
    ['Capital aportado total', DEFAULTS.LIQUIDO_INICIAL],
    ['Pico patrimonio', DEFAULTS.LIQUIDO_INICIAL]
  ]);
  liq.getRange(1,1,1,2).setFontWeight('bold'); liq.setColumnWidth(1,220);

  // GEMELO POSICIONES
  var pos = getOrCreateSheet(TABS.POS);
  pos.clear();
  pos.getRange(1,1,1,10).setValues([[
    'Ticker','Nombre','Unidades','PrecioMedio','PrecioActual','Valor€',
    'PnL€','PnL%','FechaApertura','Horizonte'
  ]]).setFontWeight('bold');
  pos.setFrozenRows(1);

  // OPERACIONES (log con trazabilidad)
  var ops = getOrCreateSheet(TABS.OPS);
  ops.clear();
  ops.getRange(1,1,1,18).setValues([[
    'Fecha','Ticker','Nombre','Acción','Unidades','Precio','Importe€','Comisión€',
    'Score','Horizonte','Motivo','Estado','PrecioCierre','PnL€','PnL%','Resultado','DíasMant.','FechaCierre'
  ]]).setFontWeight('bold');
  ops.setFrozenRows(1); ops.setColumnWidth(11,360);

  // HISTORICO
  var his = getOrCreateSheet(TABS.HIST);
  his.clear();
  his.getRange(1,1,1,7).setValues([[
    'Fecha','LíquidoGemelo€','ValorPosic.Gemelo€','PatrimGemelo€',
    'PnLRealizadoAcum€','ValorCarteraReal€','PatrimonioTotal€'
  ]]).setFontWeight('bold');
  his.setFrozenRows(1);

  // CARTERA REAL
  var rea = getOrCreateSheet(TABS.REAL);
  rea.clear();
  rea.getRange(1,1,1,11).setValues([[
    'Ticker','Nombre','Tipo','Unidades','PrecioMedio','CuotaMensual€',
    'PrecioActual','Valor€','PnL€','PnL%','Señal'
  ]]).setFontWeight('bold');
  rea.getRange(2,1,REAL_SEED.length,6).setValues(REAL_SEED);
  rea.setFrozenRows(1);

  // RECOMENDACIONES
  var rec = getOrCreateSheet(TABS.RECO);
  rec.clear();
  rec.getRange(1,1,1,3).setValues([['Ticker / Concepto','Señal','Comentario']]).setFontWeight('bold');
  rec.setFrozenRows(1); rec.setColumnWidth(3,520);

  aplicarEstilos();
  SpreadsheetApp.getActiveSpreadsheet().toast('Hoja creada y con estilo. Recarga para ver el menú "Gemelo 📈".', 'Setup OK', 6);
}

// ===================== DATOS DE MERCADO =====================
function actualizarDatos() {
  var cfg = getConfig();
  var uni = sheetData(TABS.UNIVERSO);
  var sentMap = getSentimientoMap();
  // benchmark para fuerza relativa (rentabilidad 3m de referencia)
  var bq = yahooFetch(cfg.BENCHMARK);
  var benchMom3m = bq ? computeMetrics(bq.closes, bq.price, bq.high52, bq.low52).mom3m : 0;
  var out = [];
  for (var i = 1; i < uni.length; i++) {
    var tk = String(uni[i][0]).trim(); if (!tk) continue;
    var nombre = uni[i][1], sector = uni[i][2];
    var q = yahooFetch(tk);
    if (!q) { out.push([tk, nombre, sector, 'ERROR', '', '', '', '', '', '', '', '', '', '', 'SIN DATOS', nowStr(), '']); continue; }
    var m = computeMetrics(q.closes, q.price, q.high52, q.low52);
    var sent = sentMap[sector] != null ? sentMap[sector] : 0;
    var rs = m.mom3m - benchMom3m;
    var sc = scoreStock(m, sent, cfg, rs);
    var senal = sc.score >= cfg.UMBRAL_COMPRA ? 'COMPRAR' : (sc.score < cfg.UMBRAL_VENTA ? 'EVITAR/VENDER' : 'NEUTRO');
    out.push([
      tk, nombre, sector, r2(q.price), r2(q.chgPct), r2(m.ma50), r2(m.ma200),
      r2(m.mom1m), r2(m.mom3m), r2(m.mom6m), r2(m.vol), r2(m.pos52w),
      sent, Math.round(sc.score), senal, nowStr(), r2(rs)
    ]);
    Utilities.sleep(120); // cortesía con Yahoo
  }
  var dat = getOrCreateSheet(TABS.DATOS);
  if (dat.getLastRow() > 1) dat.getRange(2,1,dat.getLastRow()-1,17).clearContent();
  if (out.length) dat.getRange(2,1,out.length,17).setValues(out);
  // ordenar por sector y luego por score desc (efecto "heatmap agrupado")
  if (out.length) dat.getRange(2,1,out.length,17).sort([{column:3, ascending:true},{column:14, ascending:false}]);
  SpreadsheetApp.getActiveSpreadsheet().toast(out.length + ' valores actualizados.', 'Datos', 4);
}

// Yahoo Finance v8 chart (sin auth, server-side). Devuelve precio, % cambio, cierres y rango 52s.
function yahooFetch(symbol) {
  try {
    var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol) + '?range=1y&interval=1d';
    var res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }
    });
    if (res.getResponseCode() !== 200) return null;
    var j = JSON.parse(res.getContentText());
    var r = j && j.chart && j.chart.result && j.chart.result[0];
    if (!r) return null;
    var meta = r.meta || {};
    var ind = r.indicators || {};
    var rawArr = (ind.quote && ind.quote[0] && ind.quote[0].close) || [];
    var adjArr = (ind.adjclose && ind.adjclose[0] && ind.adjclose[0].adjclose) || [];
    // Preferimos cierre AJUSTADO (incorpora dividendos y splits) para las señales
    var src = (adjArr && adjArr.length) ? adjArr : rawArr;
    var closes = src.filter(function(x){ return x != null; });
    if (closes.length < 30) return null;
    var price = meta.regularMarketPrice != null ? meta.regularMarketPrice : closes[closes.length-1];
    var prev = meta.chartPreviousClose != null ? meta.chartPreviousClose : closes[closes.length-2];
    var chgPct = prev ? (price/prev - 1) * 100 : 0;
    var w52 = closes.slice(-252);
    var high52 = Math.max.apply(null, w52);
    var low52  = Math.min.apply(null, w52);
    return { price: price, chgPct: chgPct, closes: closes, high52: high52, low52: low52 };
  } catch (e) { return null; }
}

function computeMetrics(closes, price, high52, low52) {
  var n = closes.length;
  function sma(k){ if (n < k) return null; var s=0; for (var i=n-k;i<n;i++) s+=closes[i]; return s/k; }
  function ret(k){ var idx=n-1-k; if (idx<0) return 0; return (price/closes[idx]-1)*100; }
  // volatilidad anualizada (últimos 63 días)
  var win = Math.min(63, n-1), rs = [];
  for (var i=n-win;i<n;i++) rs.push(closes[i]/closes[i-1]-1);
  var mean = rs.reduce(function(a,b){return a+b;},0)/rs.length;
  var varr = rs.reduce(function(a,b){return a+(b-mean)*(b-mean);},0)/rs.length;
  var vol = Math.sqrt(varr)*Math.sqrt(252)*100;
  var pos52w = (high52>low52) ? (price-low52)/(high52-low52) : 0.5;
  return {
    ma50: sma(50), ma200: sma(200),
    mom1m: ret(21), mom3m: ret(63), mom6m: ret(126),
    vol: vol, pos52w: pos52w, price: price
  };
}

// Score 0..100, transparente y ajustable por pesos. rs = exceso de rentabilidad 3m vs benchmark (%).
function scoreStock(m, sent, cfg, rs) {
  // Momentum: blend ponderado de 1m/3m/6m -> 0..100
  var blended = 0.5*m.mom1m + 0.35*m.mom3m + 0.15*m.mom6m;
  var momScore = clamp(50 + blended*2.0, 0, 100);
  // Tendencia: 0 / 50 / 100
  var trendScore = 0;
  if (m.ma50 != null && m.price > m.ma50) trendScore += 50;
  if (m.ma50 != null && m.ma200 != null && m.ma50 > m.ma200) trendScore += 50;
  // Posición en rango 52s
  var posScore = clamp(m.pos52w*100, 0, 100);
  // Sentimiento sector: -2..+2 -> 0..100
  var sentScore = clamp(50 + sent*25, 0, 100);
  // Fuerza relativa: exceso vs benchmark -> 0..100
  var rsVal = (rs==null) ? 0 : rs;
  var rsScore = clamp(50 + rsVal*2.0, 0, 100);
  var wf = cfg.W_FUERZA || 0;
  var wsum = cfg.W_MOMENTUM + cfg.W_TENDENCIA + cfg.W_POS52W + cfg.W_SENTIMIENTO + wf;
  var score = (momScore*cfg.W_MOMENTUM + trendScore*cfg.W_TENDENCIA + posScore*cfg.W_POS52W + sentScore*cfg.W_SENTIMIENTO + rsScore*wf) / wsum;
  return { score: score, mom: momScore, trend: trendScore, pos: posScore, sent: sentScore, fuerza: rsScore, m: m };
}

// ===================== MOTOR DE DECISIONES =====================
function generarDecisiones() {
  var cfg = getConfig();
  var dat = sheetData(TABS.DATOS);
  if (dat.length < 2) { SpreadsheetApp.getUi().alert('Primero pulsa "① Actualizar datos".'); return; }

  // mapa de datos por ticker
  var D = {};
  for (var i=1;i<dat.length;i++){
    var tk = String(dat[i][0]).trim(); if (!tk || dat[i][3]==='ERROR') continue;
    D[tk] = {
      nombre: dat[i][1], sector: dat[i][2], price: num(dat[i][3]), chg: num(dat[i][4]),
      ma50: num(dat[i][5]), ma200: num(dat[i][6]), mom3m: num(dat[i][8]),
      vol: num(dat[i][10]), pos52w: num(dat[i][11]), sent: num(dat[i][12]), score: num(dat[i][13])
    };
  }
  var liq = getLiquidez();
  var positions = getPositions(); // {ticker:{unidades,precioMedio,nombre,fechaApertura,horizonte}}
  var decisiones = [];

  // --- VENTAS (sobre lo que tenemos) ---
  for (var tk in positions) {
    var p = positions[tk];
    var d = D[tk];
    if (!d) continue;
    var pnlPct = (d.price/p.precioMedio - 1) * 100;
    var motivo = null;
    if (pnlPct <= cfg.STOP_LOSS_PCT) motivo = 'Stop-loss ' + r2(pnlPct) + '%';
    else if (pnlPct >= cfg.TAKE_PROFIT_PCT) motivo = 'Objetivo +' + r2(pnlPct) + '% alcanzado';
    else if (d.score < cfg.UMBRAL_VENTA) motivo = 'Score cayó a ' + Math.round(d.score) + ' (<' + cfg.UMBRAL_VENTA + ')';
    if (motivo) decisiones.push({ accion:'VENTA', tk:tk, nombre:p.nombre, unidades:p.unidades, precio:d.price, score:Math.round(d.score), horizonte:p.horizonte, motivo:motivo });
  }

  // --- COMPRAS (mejores scores que no tengamos) ---
  var equity = liq.liquido + valorPosiciones(positions, D);
  // actualizar pico de patrimonio (para el límite de caída)
  if (equity > liq.pico) { liq.pico = equity; setLiquidez(liq); }
  var capEur = equity * cfg.CAP_POSICION_PCT / 100;
  var libres = cfg.MAX_POSICIONES - Object.keys(positions).length;

  // --- FILTROS DE RIESGO (B) ---
  var riskOn = cfg.FILTRO_REGIMEN ? regimenAlcista(cfg) : true;
  var enDrawdown = (liq.pico > 0) && (equity < liq.pico * (1 + cfg.MAX_DRAWDOWN_PCT/100));
  var bloqueo = !riskOn ? ('mercado bajista (' + cfg.BENCHMARK + ' bajo su MA200)')
              : (enDrawdown ? ('límite de caída de cartera ' + cfg.MAX_DRAWDOWN_PCT + '% alcanzado') : '');
  // recuento de posiciones por sector (tope de concentración)
  var porSector = {};
  for (var ps in positions) { var sec = (D[ps] && D[ps].sector) || '?'; porSector[sec] = (porSector[sec]||0) + 1; }

  if (libres > 0 && !bloqueo) {
    var cands = [];
    for (var t in D) {
      if (positions[t]) continue;
      if (D[t].score >= cfg.UMBRAL_COMPRA && D[t].price > 0) cands.push({tk:t, d:D[t]});
    }
    cands.sort(function(a,b){ return b.d.score - a.d.score; });
    // seleccionar respetando el tope por sector
    var seleccion = [];
    for (var ci=0; ci<cands.length && seleccion.length<libres; ci++){
      var sec2 = cands[ci].d.sector || '?';
      if ((porSector[sec2]||0) >= cfg.MAX_POR_SECTOR) continue;
      porSector[sec2] = (porSector[sec2]||0) + 1;
      seleccion.push(cands[ci]);
    }
    var cajaInvertible = Math.max(0, liq.liquido - equity * cfg.BUFFER_LIQUIDEZ_PCT / 100);
    var baseEur = seleccion.length ? cajaInvertible / seleccion.length : 0;
    var medianVol = median(seleccion.map(function(c){ return c.d.vol || 25; }));
    seleccion.forEach(function(c){
      // dimensionado por volatilidad inversa, con tope por posición
      var size = baseEur * (medianVol / (c.d.vol || medianVol));
      size = Math.min(size, capEur, cajaInvertible);
      var unidades = Math.floor((size - cfg.COMISION) / c.d.price);
      if (unidades < 1) return;
      var importe = unidades * c.d.price;
      var horizonte = (c.d.ma50 && c.d.ma200 && c.d.ma50 > c.d.ma200 && c.d.vol < medianVol) ? 'Largo' : 'Corto';
      decisiones.push({ accion:'COMPRA', tk:c.tk, nombre:c.d.nombre, unidades:unidades, precio:c.d.price, score:Math.round(c.d.score), horizonte:horizonte, motivo:motivoCompra(c.d) });
      cajaInvertible -= (importe + cfg.COMISION);
    });
  }

  if (!decisiones.length) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      bloqueo ? ('Sin compras: ' + bloqueo + '. (Solo se evalúan ventas.)') : 'Sin operaciones nuevas según las reglas.',
      'Decisiones', 6);
    return;
  }

  // escribir en Operaciones
  var ops = getOrCreateSheet(TABS.OPS);
  var rows = decisiones.map(function(x){
    return [ new Date(), x.tk, x.nombre, x.accion, x.unidades, r2(x.precio),
             r2(x.unidades*x.precio), cfg.COMISION, x.score, x.horizonte, x.motivo,
             cfg.AUTO_APROBAR ? 'APROBADA' : 'PROPUESTA', '', '', '', '', '', '' ];
  });
  ops.getRange(ops.getLastRow()+1, 1, rows.length, 18).setValues(rows);

  if (cfg.AUTO_APROBAR) ejecutarAprobadas();
  else SpreadsheetApp.getActiveSpreadsheet().toast(rows.length + ' propuestas escritas. Revisa "Operaciones" y pon APROBADA para ejecutar.', 'Decisiones', 6);
}

function motivoCompra(d) {
  var parts = ['Score ' + Math.round(d.score) + '/100', 'Mom3m ' + (d.mom3m>=0?'+':'') + r2(d.mom3m) + '%'];
  if (d.ma50 && d.price > d.ma50) parts.push('Precio>MA50');
  if (d.ma50 && d.ma200 && d.ma50 > d.ma200) parts.push('MA50>MA200');
  parts.push('Pos52s ' + r2(d.pos52w));
  if (d.sent) parts.push('Sentim. ' + d.sector + ' ' + (d.sent>0?'+':'') + d.sent);
  return parts.join(' · ');
}

// ===================== EJECUCIÓN =====================
function ejecutarAprobadas() {
  var cfg = getConfig();
  var ops = getOrCreateSheet(TABS.OPS);
  var data = ops.getDataRange().getValues();
  var liq = getLiquidez();
  var positions = getPositions();
  var ejecutadas = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][11]).toUpperCase() !== 'APROBADA') continue;
    var accion = data[i][3], tk = data[i][1], nombre = data[i][2];
    var unidades = num(data[i][4]), precio = num(data[i][5]), comision = num(data[i][7]);

    if (accion === 'COMPRA') {
      var coste = unidades*precio + comision;
      if (coste > liq.liquido) { ops.getRange(i+1,12).setValue('RECHAZADA'); ops.getRange(i+1,11).setValue(data[i][10] + ' | sin liquidez'); continue; }
      liq.liquido -= coste; liq.aportado = liq.aportado; // aportado no cambia con compras
      var p = positions[tk] || { unidades:0, precioMedio:0, nombre:nombre, fechaApertura:new Date(), horizonte:data[i][9] };
      var nuevoTotal = p.unidades + unidades;
      p.precioMedio = (p.precioMedio*p.unidades + precio*unidades) / nuevoTotal;
      p.unidades = nuevoTotal; p.nombre = nombre;
      positions[tk] = p;
      ops.getRange(i+1,12).setValue('EJECUTADA');
      ejecutadas++;
    } else if (accion === 'VENTA') {
      var pos = positions[tk]; if (!pos) { ops.getRange(i+1,12).setValue('RECHAZADA'); continue; }
      var uVend = Math.min(unidades, pos.unidades);
      var ingresos = uVend*precio - comision;
      var pnl = (precio - pos.precioMedio) * uVend;   // P&L de precio (ganancias libres de impuestos)
      var pnlPct = (precio/pos.precioMedio - 1) * 100;
      var dias = Math.round((new Date() - new Date(pos.fechaApertura)) / 86400000);
      liq.liquido += ingresos;
      liq.pnlAcum += pnl;
      pos.unidades -= uVend;
      if (pos.unidades <= 0) delete positions[tk]; else positions[tk] = pos;
      ops.getRange(i+1, 12, 1, 7).setValues([[ 'EJECUTADA', r2(precio), r2(pnl), r2(pnlPct),
        (pnl>=0?'GANANCIA':'PÉRDIDA'), dias, new Date() ]]);
      ejecutadas++;
    }
  }
  setLiquidez(liq);
  writePositions(positions);
  refrescarValorPosiciones();
  actualizarPico();
  SpreadsheetApp.getActiveSpreadsheet().toast(ejecutadas + ' operaciones ejecutadas. Líquido: ' + r2(liq.liquido) + ' €', 'Ejecución', 6);
}

// ===================== CARTERA REAL + DCA =====================
function evaluarCarteraReal() {
  var cfg = getConfig();
  var rea = getOrCreateSheet(TABS.REAL);
  var data = rea.getDataRange().getValues();
  var rec = getOrCreateSheet(TABS.RECO);
  rec.clear();
  rec.getRange(1,1,1,3).setValues([['Ticker / Concepto','Señal','Comentario']]).setFontWeight('bold');
  rec.setColumnWidth(3,520);

  var recRows = [], pos52wList = [], valorTotal = 0, cuotaTotal = 0;
  for (var i = 1; i < data.length; i++) {
    var tk = String(data[i][0]).trim(); if (!tk) continue;
    var unidades = num(data[i][3]), pmedio = num(data[i][4]), cuota = num(data[i][5]);
    cuotaTotal += cuota;
    var q = yahooFetch(tk);
    if (!q) { rea.getRange(i+1,7,1,5).setValues([['ERROR','','','','sin datos']]); continue; }
    var m = computeMetrics(q.closes, q.price, q.high52, q.low52);
    var valor = unidades*q.price, pnl = pmedio ? (q.price-pmedio)*unidades : 0, pnlPct = pmedio ? (q.price/pmedio-1)*100 : 0;
    valorTotal += valor;
    var señal = 'Mantener';
    if (m.ma200 && q.price > m.ma200 && m.mom3m > 0) señal = 'Mantener/ampliar';
    else if ((m.ma200 && q.price < m.ma200) || m.mom3m < -8) señal = 'Vigilar';
    rea.getRange(i+1,7,1,5).setValues([[ r2(q.price), r2(valor), r2(pnl), r2(pnlPct), señal ]]);
    pos52wList.push(m.pos52w);
    recRows.push([tk, señal, 'Pos52s ' + r2(m.pos52w) + ' · Mom3m ' + r2(m.mom3m) + '% · ' + (m.ma200 && q.price>m.ma200 ? 'sobre MA200' : 'bajo MA200')]);
    Utilities.sleep(120);
  }

  // Análisis DCA: ¿subir cuota o aporte extraordinario?
  var avgPos = pos52wList.length ? pos52wList.reduce(function(a,b){return a+b;},0)/pos52wList.length : 0.5;
  var dcaSenal, dcaTxt;
  if (avgPos < 0.45) {
    dcaSenal = 'Aporte extraordinario (sesgo)';
    dcaTxt = 'Tu cesta cotiza lejos de máximos (pos. media 52s ' + r2(avgPos) + '). Históricamente, aportar capital en debilidad ha mejorado la rentabilidad futura esperada, aunque con MÁS varianza a corto. La aportación única ha batido al DCA ~2 de cada 3 veces en el largo plazo porque el mercado tiende a subir. Si tienes colchón de emergencia y horizonte largo, un aporte extra aquí es defendible; el DCA reduce el arrepentimiento si sigue cayendo.';
  } else if (avgPos > 0.85) {
    dcaSenal = 'Mantener cuota (sin prisa)';
    dcaTxt = 'Tu cesta cotiza cerca de máximos (pos. media 52s ' + r2(avgPos) + '). Subir mucho la cuota o meter un extra grande implica comprar caro. Mantener la aportación periódica y promediar es razonable; reserva pólvora por si hay corrección. (Recuerda: nadie predice el techo de forma fiable.)';
  } else {
    dcaSenal = 'Neutral / seguir el plan';
    dcaTxt = 'Posición media 52s ' + r2(avgPos) + ', situación intermedia. Lo más sólido es no cambiar de plan por ruido: mantén la cuota; si te sobra liquidez y el horizonte es largo, un aporte único es estadísticamente algo mejor que trocearlo, pero la diferencia es pequeña.';
  }

  if (recRows.length) rec.getRange(2,1,recRows.length,3).setValues(recRows);
  var base = recRows.length + 2;
  rec.getRange(base+1,1,3,3).setValues([
    ['— ANÁLISIS APORTACIÓN —','',''],
    ['DCA / Aporte extra', dcaSenal, dcaTxt],
    ['Cuota mensual total', r2(cuotaTotal) + ' €', 'Valor actual de la cartera real: ' + r2(valorTotal) + ' €']
  ]);
  rec.getRange(base+1,1,1,1).setFontWeight('bold');
  SpreadsheetApp.getActiveSpreadsheet().toast('Cartera real evaluada. Mira "Recomendaciones".', 'Cartera real', 5);
  return valorTotal;
}

// ===================== SNAPSHOT (trazabilidad temporal) =====================
function tomarSnapshot() {
  refrescarValorPosiciones();
  var liq = getLiquidez();
  var positions = getPositions();
  var D = datosMapPrecios();
  var valPos = valorPosiciones(positions, D);
  var patGemelo = liq.liquido + valPos;
  if (patGemelo > liq.pico) { liq.pico = patGemelo; setLiquidez(liq); }
  var valReal = valorCarteraRealRapido();
  var his = getOrCreateSheet(TABS.HIST);
  his.appendRow([ new Date(), r2(liq.liquido), r2(valPos), r2(patGemelo), r2(liq.pnlAcum), r2(valReal), r2(patGemelo + valReal) ]);
  SpreadsheetApp.getActiveSpreadsheet().toast('Snapshot guardado. Patrimonio gemelo: ' + r2(patGemelo) + ' €', 'Histórico', 5);
}

// ===================== IA OPCIONAL (overlay geopolítico) =====================
function actualizarSentimientoIA() {
  var cfg = getConfig();
  if (!cfg.USAR_IA) { SpreadsheetApp.getUi().alert('USAR_IA está en false (Config). Actívalo y pega ANTHROPIC_API_KEY para usar la IA.'); return; }
  if (!cfg.ANTHROPIC_API_KEY) { SpreadsheetApp.getUi().alert('Falta ANTHROPIC_API_KEY en Config.'); return; }
  var sectores = SENT_SEED.join(', ');
  var prompt = 'Eres analista de mercados. Busca noticias geopolíticas y macro recientes y valora el sentimiento a 1-3 meses para estos sectores bursátiles: ' + sectores +
    '. Devuelve SOLO un array JSON, sin texto ni markdown, con objetos {"sector":"...","sentimiento":<entero -2 a 2>,"nota":"<máx 12 palabras>"}.';
  try {
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': cfg.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001', max_tokens: 800,
        tools: [{ type: 'web_search_20250305', name: 'web_search' }],
        messages: [{ role: 'user', content: prompt }]
      })
    });
    var j = JSON.parse(res.getContentText());
    var txt = (j.content || []).filter(function(b){ return b.type === 'text'; }).map(function(b){ return b.text; }).join('\n');
    txt = txt.replace(/```json/gi,'').replace(/```/g,'').trim();
    var arr = JSON.parse(txt.substring(txt.indexOf('['), txt.lastIndexOf(']')+1));
    var sen = getOrCreateSheet(TABS.SENT);
    var map = {}; arr.forEach(function(o){ map[o.sector] = o; });
    var data = sen.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var s = data[i][0]; if (map[s]) {
        sen.getRange(i+1,2).setValue(clamp(Math.round(num(map[s].sentimiento)), -2, 2));
        sen.getRange(i+1,3).setValue(map[s].nota || '');
        sen.getRange(i+1,4).setValue(nowStr());
      }
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Sentimiento sectorial actualizado con IA. Reejecuta "① Actualizar datos".', 'IA', 6);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error IA: ' + e.message);
  }
}

// ===================== API JSON / JSONP (para la PWA) =====================
function doGet(e) {
  var prm = (e && e.parameter) || {};
  var action = prm.action;
  // acciones disparables desde la PWA (envueltas: el contexto web no tiene UI/toast)
  function run(fn){ try { fn(); } catch(err) {} }
  if (action === 'actualizar') run(actualizarDatos);
  if (action === 'generar')    run(generarDecisiones);
  if (action === 'ejecutar')   run(ejecutarAprobadas);
  if (action === 'evaluar')    run(evaluarCarteraReal);
  if (action === 'snapshot')   run(tomarSnapshot);

  var payload = {
    actualizado: nowStr(),
    config: { autoAprobar: getConfig().AUTO_APROBAR, benchmark: getConfig().BENCHMARK },
    liquidez: getLiquidez(),
    datos: sheetData(TABS.DATOS),
    posiciones: sheetData(TABS.POS),
    operaciones: sheetData(TABS.OPS).slice(-50),
    historico: sheetData(TABS.HIST),
    carteraReal: sheetData(TABS.REAL),
    recomendaciones: sheetData(TABS.RECO)
  };
  var json = JSON.stringify(payload);
  // JSONP: si la PWA pasa ?callback=fn, devolvemos JS y esquivamos CORS
  if (prm.callback) {
    return ContentService.createTextOutput(prm.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ===================== ESTILO / SELECTORES / PANEL =====================
var PALETA = {
  header:'#1f2937', headerFg:'#ffffff', accent:'#f59e0b',
  verde:'#16a34a', rojo:'#dc2626',
  bgVerde:'#dcfce7', bgRojo:'#fee2e2', bgAmbar:'#fef3c7', bgAzul:'#dbeafe', bgGris:'#f1f5f9', bgCard:'#f8fafc'
};
var EUR = '#,##0.00" €"';
var EURc = '[Green]#,##0.00" €";[Red]-#,##0.00" €"';
var PCTc = '[Green]0.0"%";[Red]-0.0"%"';
var NUMc = '[Green]#,##0.00;[Red]-#,##0.00';

function aplicarEstilos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // CONFIG: checkboxes para los true/false + notas
  var cfg = ss.getSheetByName(TABS.CONFIG);
  if (cfg) {
    cabecera(cfg, 2);
    cfg.getRange(2,2,Math.max(cfg.getLastRow()-1,1),1).setHorizontalAlignment('center');
    ponCheckbox(cfg, 'AUTO_APROBAR', 'Activado = ejecuta solo (autónomo total). Desactivado = te propone y tú apruebas.');
    ponCheckbox(cfg, 'USAR_IA', 'Activado = usa IA para el sentimiento geopolítico (consume tokens). Requiere ANTHROPIC_API_KEY.');
    ponCheckbox(cfg, 'FILTRO_REGIMEN', 'Activado = solo abre compras si el benchmark está por encima de su MA200 (mercado alcista).');
    bandas(cfg, 2, Math.max(cfg.getLastRow(),3), 2);
    cfg.setTabColor('#64748b');
  }

  // DATOS: heatmap (gradientes) + colores de Señal + formatos
  var dat = ss.getSheetByName(TABS.DATOS);
  if (dat) {
    cabecera(dat, 17);
    dat.setConditionalFormatRules([
      gradiente(dat.getRange('N2:N300'), 20, 50, 80),
      gradiente(dat.getRange('E2:E300'), -5, 0, 5),
      gradiente(dat.getRange('Q2:Q300'), -15, 0, 15),
      textoColor(dat.getRange('O2:O300'), 'COMPRAR', PALETA.bgVerde, PALETA.verde),
      textoColor(dat.getRange('O2:O300'), 'NEUTRO', PALETA.bgGris, '#475569'),
      textoColor(dat.getRange('O2:O300'), 'EVITAR/VENDER', PALETA.bgRojo, PALETA.rojo)
    ]);
    fmt(dat, 'D2:G300', '#,##0.00');
    fmt(dat, 'E2:E300', '0.0"%"');
    fmt(dat, 'H2:K300', '0.0"%"');
    fmt(dat, 'L2:L300', '0.00');
    fmt(dat, 'Q2:Q300', '0.0"%"');
    bandas(dat, 2, 300, 17);
    dat.setTabColor('#2563eb');
  }

  // OPERACIONES: dropdown de Estado + colores + formatos
  var ops = ss.getSheetByName(TABS.OPS);
  if (ops) {
    cabecera(ops, 18);
    desplegable(ops.getRange('L2:L1000'), ['PROPUESTA','APROBADA','EJECUTADA','RECHAZADA']);
    ops.setConditionalFormatRules([
      textoColor(ops.getRange('L2:L1000'), 'PROPUESTA', PALETA.bgAmbar, '#92400e'),
      textoColor(ops.getRange('L2:L1000'), 'APROBADA', PALETA.bgAzul, '#1e40af'),
      textoColor(ops.getRange('L2:L1000'), 'EJECUTADA', PALETA.bgVerde, PALETA.verde),
      textoColor(ops.getRange('L2:L1000'), 'RECHAZADA', PALETA.bgRojo, PALETA.rojo),
      textoColor(ops.getRange('D2:D1000'), 'COMPRA', PALETA.bgVerde, PALETA.verde),
      textoColor(ops.getRange('D2:D1000'), 'VENTA', PALETA.bgRojo, PALETA.rojo),
      textoColor(ops.getRange('P2:P1000'), 'GANANCIA', PALETA.bgVerde, PALETA.verde),
      textoColor(ops.getRange('P2:P1000'), 'PÉRDIDA', PALETA.bgRojo, PALETA.rojo)
    ]);
    fmt(ops, 'F2:H1000', '#,##0.00');
    fmt(ops, 'M2:M1000', '#,##0.00');
    fmt(ops, 'N2:N1000', NUMc);
    fmt(ops, 'O2:O1000', PCTc);
    ops.setColumnWidth(11, 360);
    bandas(ops, 2, 300, 18);
    ops.setTabColor('#16a34a');
  }

  // GEMELO POSICIONES
  var pos = ss.getSheetByName(TABS.POS);
  if (pos) {
    cabecera(pos, 10);
    fmt(pos, 'D2:F1000', '#,##0.00');
    fmt(pos, 'G2:G1000', NUMc);
    fmt(pos, 'H2:H1000', PCTc);
    bandas(pos, 2, 200, 10);
    pos.setTabColor('#0d9488');
  }

  // GEMELO LIQUIDEZ
  var liq = ss.getSheetByName(TABS.LIQ);
  if (liq) {
    cabecera(liq, 2);
    liq.getRange('B2:B5').setNumberFormat(EUR).setFontSize(13).setFontWeight('bold');
    liq.setTabColor('#f59e0b');
  }

  // SENTIMIENTO SECTOR: dropdown -2..2 + gradiente
  var sen = ss.getSheetByName(TABS.SENT);
  if (sen) {
    cabecera(sen, 4);
    desplegable(sen.getRange('B2:B1000'), [-2,-1,0,1,2]);
    sen.setConditionalFormatRules([ gradiente(sen.getRange('B2:B1000'), -2, 0, 2) ]);
    sen.getRange('B2:B1000').setHorizontalAlignment('center');
    sen.setColumnWidth(3, 320);
    bandas(sen, 2, Math.max(sen.getLastRow(),4), 4);
    sen.setTabColor('#7c3aed');
  }

  // CARTERA REAL: dropdown de Tipo + colores + formatos
  var rea = ss.getSheetByName(TABS.REAL);
  if (rea) {
    cabecera(rea, 11);
    desplegable(rea.getRange('C2:C100'), ['Acción','ETF','Fondo','Otro']);
    rea.setConditionalFormatRules([
      textoColor(rea.getRange('K2:K100'), 'Mantener/ampliar', PALETA.bgVerde, PALETA.verde),
      textoColor(rea.getRange('K2:K100'), 'Vigilar', PALETA.bgAmbar, '#92400e'),
      textoColor(rea.getRange('K2:K100'), 'Mantener', PALETA.bgGris, '#475569')
    ]);
    fmt(rea, 'E2:E100', '#,##0.00');
    fmt(rea, 'F2:F100', EUR);
    fmt(rea, 'G2:G100', '#,##0.00');
    fmt(rea, 'H2:H100', EUR);
    fmt(rea, 'I2:I100', NUMc);
    fmt(rea, 'J2:J100', PCTc);
    bandas(rea, 2, 100, 11);
    rea.setTabColor('#0891b2');
  }

  // HISTORICO
  var his = ss.getSheetByName(TABS.HIST);
  if (his) {
    cabecera(his, 7);
    fmt(his, 'B2:G2000', EUR);
    bandas(his, 2, 400, 7);
    his.setTabColor('#475569');
  }

  // RECOMENDACIONES
  var rec = ss.getSheetByName(TABS.RECO);
  if (rec) {
    cabecera(rec, 3);
    rec.setColumnWidth(3, 520);
    rec.getRange('C2:C1000').setWrap(true);
    rec.setTabColor('#ea580c');
  }

  crearPanel();
  SpreadsheetApp.getActiveSpreadsheet().toast('Estilos y selectores aplicados.', 'Listo 🎨', 4);
}

// ---- helpers de estilo ----
function cabecera(sh, nCols){
  sh.getRange(1,1,1,nCols).setBackground(PALETA.header).setFontColor(PALETA.headerFg)
    .setFontWeight('bold').setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setFrozenRows(1); sh.setRowHeight(1, 30);
}
function bandas(sh, fila, hasta, nCols){
  try { sh.getBandings().forEach(function(b){ b.remove(); }); } catch(e){}
  if (hasta < fila) return;
  try { sh.getRange(fila,1,hasta-fila+1, nCols)
          .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch(e){}
}
function fmt(sh, a1, format){ sh.getRange(a1).setNumberFormat(format); }
function desplegable(range, lista){
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(lista, true).setAllowInvalid(true).build());
}
function gradiente(range, min, mid, max){
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue('#1a9850', SpreadsheetApp.InterpolationType.NUMBER, String(max))
    .setGradientMidpointWithValue('#ffffbf', SpreadsheetApp.InterpolationType.NUMBER, String(mid))
    .setGradientMinpointWithValue('#d73027', SpreadsheetApp.InterpolationType.NUMBER, String(min))
    .setRanges([range]).build();
}
function textoColor(range, texto, bg, fg){
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(texto).setBackground(bg).setFontColor(fg).setRanges([range]).build();
}
function ponCheckbox(sh, clave, nota){
  var data = sh.getRange(1,1,sh.getLastRow(),1).getValues();
  for (var i=0;i<data.length;i++){
    if (String(data[i][0]).trim() === clave){
      var cell = sh.getRange(i+1,2);
      var v = String(cell.getValue()).toLowerCase();
      var on = (cell.getValue()===true || v==='true' || v==='sí' || v==='si' || v==='1');
      cell.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
      cell.setValue(on);
      if (nota) sh.getRange(i+1,1).setNote(nota);
      return;
    }
  }
}

// ===================== PANEL (dashboard) =====================
function crearPanel(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = ss.getSheetByName('Panel') || ss.insertSheet('Panel');
  p.clear();
  p.getRange(1,1,40,10).clearDataValidations();
  p.setHiddenGridlines(true);

  p.setColumnWidth(1, 24);
  [2,4,6].forEach(function(c){ p.setColumnWidth(c, 175); });
  [3,5,7].forEach(function(c){ p.setColumnWidth(c, 24); });

  // título
  p.getRange('B2:F2').merge().setValue('📈  GEMELO DIGITAL DE MERCADO')
    .setBackground(PALETA.header).setFontColor('#ffffff').setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  p.setRowHeight(2, 38);
  p.getRange('B3:F3').merge().setValue('Panel de control · usa el menú "Gemelo 📈" para actualizar')
    .setFontColor('#64748b').setHorizontalAlignment('center').setFontSize(9);

  var L="'GemeloLiquidez'", G="'GemeloPosiciones'", R="'CarteraReal'";
  tarjeta(p,'B5','B6','LÍQUIDO DISPONIBLE', "="+L+"!B2", EUR, PALETA.accent);
  tarjeta(p,'D5','D6','P&L REALIZADO', "="+L+"!B3", EURc, '#0f172a');
  tarjeta(p,'F5','F6','PATRIMONIO GEMELO', "="+L+"!B2 + SUM("+G+"!F2:F)", EUR, '#0f172a');
  tarjeta(p,'B9','B10','RENTABILIDAD GEMELO', "=IF("+L+"!B4=0; 0; "+L+"!B3/"+L+"!B4)", '[Green]0.0%;[Red]-0.0%', '#0f172a');
  tarjeta(p,'D9','D10','Nº POSICIONES', "=COUNTA("+G+"!A2:A)", '0', '#0f172a');
  tarjeta(p,'F9','F10','VALOR CARTERA REAL', "=SUM("+R+"!H2:H)", EUR, '#0f172a');

  // curva de patrimonio
  p.getRange('B12:F12').merge().setValue('Evolución del patrimonio del gemelo')
    .setFontWeight('bold').setFontColor('#334155').setFontSize(11);
  p.getRange('B13:F16').merge().setFormula('=SPARKLINE(Historico!D2:D)').setBackground(PALETA.bgCard);
  p.getRange('B13:F16').setBorder(true,true,true,true,false,false,'#e2e8f0',SpreadsheetApp.BorderStyle.SOLID);

  p.getRange('B18:F18').merge().setValue('Flujo: ① Actualizar datos → ② Generar decisiones → aprobar en Operaciones → ③ Ejecutar → ⑤ Snapshot')
    .setFontColor('#64748b').setFontSize(9);

  ss.setActiveSheet(p); ss.moveActiveSheet(1);
}
function tarjeta(sh, celdaLabel, celdaValor, label, formula, format, colorValor){
  var lab = sh.getRange(celdaLabel);
  lab.setValue(label).setFontSize(8).setFontColor('#64748b').setFontWeight('bold')
     .setBackground(PALETA.bgCard).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(lab.getRow(), 22);
  var val = sh.getRange(celdaValor);
  val.setFormula(formula).setNumberFormat(format).setFontSize(18).setFontWeight('bold')
     .setFontColor(colorValor).setBackground(PALETA.bgCard)
     .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(val.getRow(), 36);
  sh.getRange(lab.getRow(), lab.getColumn(), 2, 1).setBorder(true,true,true,true,false,false,'#e2e8f0',SpreadsheetApp.BorderStyle.SOLID);
}

// ===================== BACKTEST / BENCHMARK =====================
// Histórico diario de N años (epoch + cierres). Reutiliza la misma lógica de score que el motor en vivo.
function yahooHist(symbol, anos){
  try{
    var url='https://query1.finance.yahoo.com/v8/finance/chart/'+encodeURIComponent(symbol)+'?range='+anos+'y&interval=1d';
    var res=UrlFetchApp.fetch(url,{muteHttpExceptions:true,headers:{'User-Agent':'Mozilla/5.0','Accept':'application/json'}});
    if(res.getResponseCode()!==200) return null;
    var r=JSON.parse(res.getContentText()).chart.result[0];
    var ts=r.timestamp||[], ind=r.indicators||{};
    var adj=(ind.adjclose&&ind.adjclose[0]&&ind.adjclose[0].adjclose)||[];
    var raw=(ind.quote&&ind.quote[0]&&ind.quote[0].close)||[];
    var cl=(adj&&adj.length)?adj:raw;
    var dates=[], closes=[];
    for(var i=0;i<ts.length;i++){ if(cl[i]!=null){ dates.push(ts[i]); closes.push(cl[i]); } }
    return {ts:dates, closes:closes};
  }catch(e){ return null; }
}

function backtest(){
  var ui=SpreadsheetApp.getUi(), cfg=getConfig();
  var anos=Math.max(1, Math.round(cfg.BACKTEST_ANOS||2));
  var step=Math.max(1, Math.round(cfg.BACKTEST_REBALANCEO_DIAS||5));
  SpreadsheetApp.getActiveSpreadsheet().toast('Descargando histórico… puede tardar ~1 min.', 'Backtest', 12);

  var bench=yahooHist(cfg.BENCHMARK, anos);
  if(!bench || bench.closes.length<210){ ui.alert('No pude descargar el benchmark "'+cfg.BENCHMARK+'". Revisa el ticker en Config.'); return; }

  var uni=sheetData(TABS.UNIVERSO), S={};
  for(var i=1;i<uni.length;i++){
    var tk=String(uni[i][0]).trim(); if(!tk) continue;
    var h=yahooHist(tk, anos);
    if(h && h.closes.length>210) S[tk]={ts:h.ts, closes:h.closes, sector:uni[i][2], ptr:0};
    Utilities.sleep(120);
  }
  var tickers=Object.keys(S);
  if(!tickers.length){ ui.alert('Sin datos de universo suficientes para el backtest.'); return; }

  var cash=cfg.LIQUIDO_INICIAL, pos={}, cerrados=0, ganados=0, curva=[], peak=cfg.LIQUIDO_INICIAL;
  var start=200, bench0=bench.closes[start];
  function px(tk){ return S[tk].closes[S[tk].ptr]; }

  for(var bi=start; bi<bench.closes.length; bi+=step){
    var t=bench.ts[bi];
    for(var x=0;x<tickers.length;x++){ var s=S[tickers[x]]; while(s.ptr+1<s.ts.length && s.ts[s.ptr+1]<=t) s.ptr++; }

    var mtm=0; for(var p in pos) mtm += pos[p].u*px(p);
    var equity=cash+mtm;
    if(equity>peak) peak=equity;
    var enDD = equity < peak*(1+cfg.MAX_DRAWDOWN_PCT/100);
    var ma200b=0; for(var z=bi-199; z<=bi; z++) ma200b+=bench.closes[z]; ma200b/=200;
    var riskOn = cfg.FILTRO_REGIMEN ? (bench.closes[bi] > ma200b) : true;
    var benchMom3m = (bench.closes[bi]/bench.closes[bi-63]-1)*100;

    // scores en esta fecha (mismo cálculo que en vivo, sin overlay IA)
    var sc={};
    for(var y=0;y<tickers.length;y++){
      var tk=tickers[y], s2=S[tk]; if(s2.ptr<200) continue;
      var sub=s2.closes.slice(0, s2.ptr+1);
      var w=sub.slice(-252), hi=Math.max.apply(null,w), lo=Math.min.apply(null,w);
      var m=computeMetrics(sub, sub[sub.length-1], hi, lo);
      sc[tk]={score:scoreStock(m,0,cfg, m.mom3m-benchMom3m).score, price:sub[sub.length-1], vol:m.vol};
    }

    // salidas
    for(var p2 in pos){
      var d=sc[p2]; if(!d) continue;
      var pnlPct=(d.price/pos[p2].pm-1)*100;
      if(pnlPct<=cfg.STOP_LOSS_PCT || pnlPct>=cfg.TAKE_PROFIT_PCT || d.score<cfg.UMBRAL_VENTA){
        cash += pos[p2].u*d.price - cfg.COMISION;
        cerrados++; if((d.price-pos[p2].pm)>0) ganados++;
        delete pos[p2];
      }
    }

    // entradas (con filtros de riesgo B: régimen + drawdown + tope por sector)
    var libres=cfg.MAX_POSICIONES-Object.keys(pos).length;
    if(libres>0 && riskOn && !enDD){
      var porSec={}; for(var ph in pos){ var se=S[ph].sector||'?'; porSec[se]=(porSec[se]||0)+1; }
      var cand=[]; for(var tk2 in sc){ if(!pos[tk2] && sc[tk2].score>=cfg.UMBRAL_COMPRA) cand.push(tk2); }
      cand.sort(function(a,b){return sc[b].score-sc[a].score;});
      var sel=[];
      for(var k=0;k<cand.length && sel.length<libres;k++){
        var se2=S[cand[k]].sector||'?';
        if((porSec[se2]||0)>=cfg.MAX_POR_SECTOR) continue;
        porSec[se2]=(porSec[se2]||0)+1; sel.push(cand[k]);
      }
      var capEur=equity*cfg.CAP_POSICION_PCT/100;
      var caja=Math.max(0, cash-equity*cfg.BUFFER_LIQUIDEZ_PCT/100);
      var base=sel.length?caja/sel.length:0;
      var mv=median(sel.map(function(c){return sc[c].vol||25;}));
      for(var c2=0;c2<sel.length;c2++){
        var tk3=sel[c2], d3=sc[tk3];
        var size=Math.min(base*(mv/(d3.vol||mv)), capEur, caja);
        var u=Math.floor((size-cfg.COMISION)/d3.price); if(u<1) continue;
        cash -= (u*d3.price+cfg.COMISION); caja -= (u*d3.price+cfg.COMISION);
        pos[tk3]={u:u, pm:d3.price};
      }
    }

    var mtm2=0; for(var p3 in pos) mtm2+=pos[p3].u*px(p3);
    curva.push({t:t, motor:cash+mtm2, bench:cfg.LIQUIDO_INICIAL*bench.closes[bi]/bench0});
  }

  var last=curva.length?curva[curva.length-1].motor:cfg.LIQUIDO_INICIAL;
  var R={
    motorRet:(last/cfg.LIQUIDO_INICIAL-1)*100,
    benchRet:(bench.closes[bench.closes.length-1]/bench0-1)*100,
    ddMotor:maxDD(curva.map(function(p){return p.motor;})),
    ddBench:maxDD(curva.map(function(p){return p.bench;})),
    winRate:cerrados?(ganados/cerrados*100):0,
    cerrados:cerrados, last:last, anos:anos, step:step
  };
  escribirBacktest(cfg, curva, R);
}

function maxDD(arr){ var peak=-Infinity, dd=0; for(var i=0;i<arr.length;i++){ if(arr[i]>peak) peak=arr[i]; var d=peak>0?(peak-arr[i])/peak*100:0; if(d>dd) dd=d; } return dd; }

function escribirBacktest(cfg, curva, R){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var bt=ss.getSheetByName('Backtest')||ss.insertSheet('Backtest');
  bt.clear(); bt.getCharts().forEach(function(c){ bt.removeChart(c); });
  bt.setHiddenGridlines(true);
  bt.getRange('A1').setValue('📊 BACKTEST · motor vs '+cfg.BENCHMARK).setFontSize(15).setFontWeight('bold');
  bt.getRange('A2').setValue('Últimos '+R.anos+' año(s) · rebalanceo '+R.step+'d · comisión '+cfg.COMISION+' € · régimen:'+(cfg.FILTRO_REGIMEN?'ON':'OFF')+' · máx '+cfg.MAX_POR_SECTOR+'/sector · freno drawdown '+cfg.MAX_DRAWDOWN_PCT+'% · SIN deslizamiento ni overlay IA · pasado ≠ futuro.')
    .setFontColor('#64748b').setFontSize(9);

  var resumen=[
    ['Métrica','Motor','Benchmark'],
    ['Rentabilidad total', R.motorRet/100, R.benchRet/100],
    ['Caída máxima (drawdown)', -R.ddMotor/100, -R.ddBench/100],
    ['Patrimonio final (€)', R.last, cfg.LIQUIDO_INICIAL*(1+R.benchRet/100)],
    ['% aciertos (cerradas)', R.winRate/100, ''],
    ['Nº operaciones cerradas', R.cerrados, '']
  ];
  bt.getRange(4,1,resumen.length,3).setValues(resumen);
  bt.getRange(4,1,1,3).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  bt.getRange(5,2,2,2).setNumberFormat('[Green]0.0%;[Red]-0.0%');
  bt.getRange(7,2,1,2).setNumberFormat('#,##0.00" €"');
  bt.getRange(8,2,1,1).setNumberFormat('0.0%');
  bt.setColumnWidth(1,260);

  var sRow=12;
  bt.getRange(sRow,1,1,3).setValues([['Fecha','Motor €','Benchmark €']]).setFontWeight('bold');
  var rows=curva.map(function(p){ return [new Date(p.t*1000), Math.round(p.motor*100)/100, Math.round(p.bench*100)/100]; });
  if(rows.length){
    bt.getRange(sRow+1,1,rows.length,3).setValues(rows);
    bt.getRange(sRow+1,1,rows.length,1).setNumberFormat('yyyy-mm-dd');
    var chart=bt.newChart().asLineChart()
      .addRange(bt.getRange(sRow,1,rows.length+1,3)).setNumHeaders(1)
      .setPosition(4,5,0,0)
      .setOption('title','Motor vs Benchmark (€)').setOption('legend',{position:'bottom'})
      .setOption('width',520).setOption('height',300).build();
    bt.insertChart(chart);
  }
  bt.setTabColor('#dc2626');
  ss.setActiveSheet(bt);
  SpreadsheetApp.getActiveSpreadsheet().toast('Backtest listo · Motor '+R.motorRet.toFixed(1)+'% vs '+cfg.BENCHMARK+' '+R.benchRet.toFixed(1)+'%.', 'Backtest', 8);
}

// ===================== HELPERS =====================
function getOrCreateSheet(name){ var ss=SpreadsheetApp.getActiveSpreadsheet(); return ss.getSheetByName(name) || ss.insertSheet(name); }
function sheetData(name){ var sh=getOrCreateSheet(name); return sh.getDataRange().getValues(); }
function nowStr(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function r2(x){ return (x==null||isNaN(x))?'':Math.round(x*100)/100; }
function num(x){ var n=Number(x); return isNaN(n)?0:n; }
function clamp(x,a,b){ return Math.max(a, Math.min(b, x)); }
function median(a){ if(!a.length) return 0; var s=a.slice().sort(function(x,y){return x-y;}); var m=Math.floor(s.length/2); return s.length%2?s[m]:(s[m-1]+s[m])/2; }

function getConfig(){
  var sh=getOrCreateSheet(TABS.CONFIG); var data=sh.getDataRange().getValues(); var cfg={};
  for (var k in DEFAULTS) cfg[k]=DEFAULTS[k];
  for (var i=1;i<data.length;i++){
    var key=data[i][0]; if(!key||!(key in DEFAULTS)) continue; var val=data[i][1]; var def=DEFAULTS[key];
    if (typeof def==='boolean'){ var s=String(val).toLowerCase(); cfg[key]=(val===true||s==='true'||s==='sí'||s==='si'||s==='1'); }
    else if (typeof def==='number'){ cfg[key]=(val===''||val==null)?def:Number(val); }
    else cfg[key]=(val==null?'':String(val));
  }
  return cfg;
}

function getSentimientoMap(){ var d=sheetData(TABS.SENT); var m={}; for(var i=1;i<d.length;i++){ if(d[i][0]) m[d[i][0]]=num(d[i][1]); } return m; }

function getLiquidez(){
  var d=sheetData(TABS.LIQ);
  return { liquido:num(d[1]&&d[1][1]), pnlAcum:num(d[2]&&d[2][1]), aportado:num(d[3]&&d[3][1]), pico:num(d[4]&&d[4][1]) };
}
function setLiquidez(l){
  var sh=getOrCreateSheet(TABS.LIQ);
  sh.getRange(2,2).setValue(r2(l.liquido)); sh.getRange(3,2).setValue(r2(l.pnlAcum)); sh.getRange(4,2).setValue(r2(l.aportado));
  if (l.pico!=null) sh.getRange(5,2).setValue(r2(l.pico));
}

function getPositions(){
  var d=sheetData(TABS.POS); var p={};
  for (var i=1;i<d.length;i++){ var tk=String(d[i][0]).trim(); if(!tk) continue;
    p[tk]={ nombre:d[i][1], unidades:num(d[i][2]), precioMedio:num(d[i][3]), fechaApertura:d[i][8]||new Date(), horizonte:d[i][9]||'Corto' }; }
  return p;
}
function writePositions(p){
  var sh=getOrCreateSheet(TABS.POS);
  if (sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,10).clearContent();
  var rows=[]; for (var tk in p){ var x=p[tk]; rows.push([tk,x.nombre,x.unidades,r2(x.precioMedio),'','','','',x.fechaApertura,x.horizonte]); }
  if (rows.length) sh.getRange(2,1,rows.length,10).setValues(rows);
}

function datosMapPrecios(){ var d=sheetData(TABS.DATOS); var m={}; for(var i=1;i<d.length;i++){ if(d[i][0]&&d[i][3]!=='ERROR') m[d[i][0]]={price:num(d[i][3])}; } return m; }
function valorPosiciones(p,D){ var v=0; for(var tk in p){ var px=(D[tk]&&D[tk].price)||p[tk].precioMedio; v+=p[tk].unidades*px; } return v; }

// ¿Mercado alcista? benchmark por encima de su MA200. Si falla la descarga, no bloquea (devuelve true).
function regimenAlcista(cfg){
  var b=yahooFetch(cfg.BENCHMARK); if(!b) return true;
  var m=computeMetrics(b.closes, b.price, b.high52, b.low52);
  return (m.ma200==null) ? true : (b.price > m.ma200);
}
// Actualiza el pico de patrimonio del gemelo (para el límite de drawdown).
function actualizarPico(){
  var liq=getLiquidez(), positions=getPositions(), D=datosMapPrecios();
  var eq=liq.liquido + valorPosiciones(positions, D);
  if (eq > liq.pico) { liq.pico=eq; setLiquidez(liq); }
}

function refrescarValorPosiciones(){
  var sh=getOrCreateSheet(TABS.POS); var d=sh.getDataRange().getValues(); var D=datosMapPrecios();
  for (var i=1;i<d.length;i++){ var tk=String(d[i][0]).trim(); if(!tk) continue;
    var px=(D[tk]&&D[tk].price)||num(d[i][3]); var u=num(d[i][2]); var pm=num(d[i][3]);
    var val=u*px, pnl=(px-pm)*u, pnlPct=pm?(px/pm-1)*100:0;
    sh.getRange(i+1,5,1,4).setValues([[r2(px),r2(val),r2(pnl),r2(pnlPct)]]); }
}

function valorCarteraRealRapido(){
  var d=sheetData(TABS.REAL); var v=0;
  for (var i=1;i<d.length;i++){ if(!d[i][0]) continue; v+=num(d[i][7]); } // usa Valor€ ya calculado
  return v;
}
