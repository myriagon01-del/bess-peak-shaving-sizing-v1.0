'use strict';
(function () {
  /* =====================================================================
     BESS Peak-Shaving Sizing - clean workbook-equivalent engine
     Rebuilt from the CompositePeak_ModelConfig source (v.260728).
     ===================================================================== */

  /* ------------------------- i18n ------------------------- */
  var STR = {
    en: {
      appTitle: 'BESS Peak-Shaving Sizing',
      version: 'v1.0.0',
      subtitle: 'Battery sizing from real interval data \u00b7 peak-demand window',
      brandName: 'BESS Sizing Tool',
      brandSubline: 'Peak-shaving sizing for C&I',
      tabSizing: 'Analysis & BESS Sizing',
      tabLibrary: 'Battery & PCS Library',
      chartTitle: 'Load profile & selected governing day',
      chartViewLabel: 'Chart view',
      chartViewHighest: 'Highest demand day',
      chartViewGoverning: 'Governing Energy Day',
      highestMdLabel: 'Highest demand',
      governingEnergyLabel: 'Governing energy day',
      legendLoad: 'Load (before PV)',
      legendTarget: 'Shave-to target',
      legendWindow: 'Sizing window',
      legendBattery: 'Battery discharge',
      legendGridAfter: 'Grid after BESS',
      chartNote: 'After upload, the chart shows the actual day with the highest required energy.',
      loadTitle: 'Load Data (before-PV basis)',
      loadLabel: 'Upload load profile (.xlsx) or CSV',
      csvTemplateBtn: 'Download CSV template',
      mapAdjustBtn: 'Adjust mapping',
      uploadStatusNoFile: 'No file selected.',
      upLoading: 'Reading file\u2026',
      upLowConfidence: 'Automatic detection confidence is low \u2014 please confirm the mapping.',
      upMapped: 'Loaded {n} intervals \u00b7 {unit} \u00b7 {interval} min \u00b7 {days} days',
      upNoData: 'No usable rows found.',
      upError: 'Could not read file: {err}',
      mappingTitle: 'Confirm column mapping',
      mappingSheet: 'Sheet',
      mappingHeaderRow: 'Header row',
      mappingTsMode: 'Timestamp columns',
      mappingTsModeCombined: 'Combined date + time',
      mappingTsModeSplit: 'Separate date + time',
      mappingTsCol: 'Timestamp column',
      mappingDateCol: 'Date column',
      mappingTimeCol: 'Time column',
      mappingLoadCol: 'Load column',
      mappingPvCol: 'PV generation column (optional)',
      mappingExportCol: 'Grid export column (optional)',
      mappingUnit: 'Unit',
      mappingInterval: 'Interval (minutes)',
      mappingApply: 'Apply mapping',
      mappingCancel: 'Cancel',
      mappingNone: '(none)',
      controlsTitle: 'Sizing controls',
      strategyLabel: 'Sizing strategy',
      strategyStatic: 'Static Threshold',
      strategyFlatTop: 'Dynamic Discharge',
      strategyConst: 'Constant Discharge',
      eduStatic: 'Static threshold shaves every interval above the target within the sizing window.',
      eduFlatTop: 'Dynamic Discharge follows the entered shave-to target across every complete actual day. Use Auto-calculate to find the lowest sustainable target for the current pack count.',
      eduConst: 'Constant discharge applies a fixed discharge power across the sizing window; required energy = discharge kW \u00d7 window hours.',
      targetLabel: 'Shave-to Target Grid Load (kW)',
      autoTargetBtn: 'Auto-calculate sustainable target',
      autoTargetUsing: 'Lowest sustainable target using {n} currently recommended pack(s): {target} kW.',
      autoTargetUnavailable: 'Upload valid load data and calculate a recommendation first.',
      constKwLabel: 'Discharge power (kW)',
      reserveLabel: 'Reserve margin (%)',
      demandIntervalLabel: 'Demand interval (minutes)',
      modelLabel: 'Battery model',
      designYearLabel: 'Design sizing year',
      designYearNote: 'Energy capacity is derated to Year 15. Power is not degradation-derated.',
      detailUsable: 'Usable per pack',
      detailUsableYear1: 'Usable Year 1',
      detailUsableYear15: 'Usable Year 15',
      detailSOH15: 'SOH Year 15',
      detailGross: 'Gross per pack',
      detailDoD: 'DoD',
      detailRte: 'RTE',
      detailAux: 'Aux',
      detailPackPower: 'Pack rated power',
      detailPcsPower: 'PCS rated power',
      pvModeLabel: 'PV basis',
      pvModeNone: 'No PV (grid import = before-PV load)',
      pvModeDetected: 'Use PV/export columns',
      pvModeEstimated: 'Estimated monthly PV (preliminary)',
      pvMonthlyLabel: 'Monthly PV energy (kWh)',
      pvPeakKwLabel: 'PV peak power (kW)',
      pvWarning: 'Preliminary estimate \u2014 not bankable sizing',
      sizingWindowLabel: 'Sizing window',
      mdIntervalNote: 'Demand interval (minutes) — the interval over which peak demand is measured.',
      controlWindowLabel: 'Operational control window',
      ctrlStartLabel: 'Start',
      ctrlEndLabel: 'End',
      recoTitle: 'Recommended BESS Sizing',
      noData: 'Upload load data to size the battery.',
      recoGoverning: 'Governing day',
      recoRequiredEnergy: 'Required energy',
      recoEnergyPacks: 'Energy packs',
      recoNameplate: 'Nameplate (gross)',
      recoUsable: 'Usable capacity',
      recoYear15Usable: 'Year-15 usable capacity',
      recoPower: 'Required power',
      recoPcs: 'PCS quantity',
      recoPowerPacks: 'Power-constrained packs',
      warnPower: 'Power requires {powerPacks} packs ({power} kW) but energy requires only {energyPacks}. Review the model C-rate/PCS or install the higher pack count.',
      applyReco: 'Apply Recommended Sizing',
      applied: 'Applied: usable capacity = {usable} kWh.',
      footNote: 'Energy sizing uses every complete actual day within the sizing window; the day with the highest required energy governs. Energy packs = ceil(required energy \u00d7 (1 + reserve) \u00f7 usable per pack). Power and PCS are sized separately and never silently replace the energy result.',
      tabDischarge: 'Discharge Verification',
      dvTitle: 'Discharge Verification',
      dvNote: 'Simulates the sized BESS discharging over every complete day to confirm it can shave the target without running out of energy.',
      dvMdBefore: 'Peak before',
      dvMdAfter: 'Peak after',
      dvMdReduction: 'Peak reduction',
      dvEnergy: 'Energy discharged (worst day)',
      dvMinSoc: 'Min usable energy remaining',
      dvCycleDepth: 'Cycle depth',
      dvSufficient: 'Sufficient \u2014 sized capacity discharges and shaves the target on every complete day.',
      dvInsufficient: 'Insufficient \u2014 battery depletes or the peak is not fully shaved ({day}). Review capacity / power / target.',
      libTitle: 'Battery & PCS Library',
      libNew: 'New model',
      libSave: 'Save model',
      libDuplicate: 'Duplicate',
      libExcelDownload: 'Download Excel Template',
      libExcelImport: 'Upload Excel Library',
      libReset: 'Reset',
      libDelete: 'Delete',
      libUse: 'Use',
      libEdit: 'Edit',
      libEditing: 'Editing: {name}',
      libColName: 'Name',
      libColBrand: 'Brand',
      libColGross: 'Gross kWh',
      libColDoD: 'DoD %',
      libColRte: 'RTE %',
      libColAux: 'Aux %',
      libColUsable: 'Usable/pack',
      libColSOH15: 'SOH Y15',
      libColUsable15: 'Usable Y15',
      libColPower: 'Pack kW',
      libColPcs: 'PCS kW',
      libColActions: 'Actions',
      mName: 'Name',
      mBrand: 'Brand',
      mGross: 'Gross pack kWh',
      mSocMin: 'Min SOC %',
      mSocMax: 'Max SOC %',
      mRte: 'RTE %',
      mAux: 'Auxiliary factor %',
      mPackPower: 'Pack rated power kW',
      mPcsPower: 'PCS rated power kW',
      mFirstDeg: 'First-year degradation %',
      mAnnualDeg: 'Annual degradation, years 2-15 %',
      mRteDefinition: 'RTE definition',
      mDegNotes: 'Degradation / reserve notes',
      mSrcNotes: 'Source notes',
      helperDoD: 'DoD is calculated as Max SOC - Min SOC.',
      helperAux: 'Auxiliary retention factor: 98% means 2% auxiliary loss.',
      libSaved: 'Model saved.',
      libResetDone: 'Library reset to default models.',
      libExcelImported: 'Imported {n} model(s) from Excel.',
      libExcelMissingHeaders: 'Excel import rejected: required header(s) missing: {headers}.',
      libExcelRowErrors: 'Excel import rejected: {n} invalid row(s). {detail}',
      valModelGross: 'Gross pack kWh must be greater than 0.',
      valModelSocMin: 'Min SOC % must be between 0 and 100.',
      valModelSocMax: 'Max SOC % must be between 0 and 100.',
      valModelSocOrder: 'Min SOC must be lower than Max SOC.',
      valModelRte: 'RTE % must be greater than 0 and at most 100.',
      valModelAux: 'Aux % must be greater than 0 and at most 100.',
      valModelPackPower: 'Pack rated power (kW) must be greater than 0.',
      valModelPcsPower: 'PCS rated power (kW) must be greater than 0.',
      valModelFirstDeg: 'First-year degradation must be at least 0% and below 100%.',
      valModelAnnualDeg: 'Annual degradation must be at least 0% and below 100%.',
      opReadinessOk: 'Operational window covers the sizing window.',
      opReadinessWarn: 'Operational window does not cover the full sizing window.',
      opReadinessErr: 'Operational window start must be earlier than end.',
      valTitle: 'Validation summary',
      valRaw: 'Rows read',
      valParsed: 'Rows parsed',
      valInvalid: 'Invalid / negative load rows excluded',
      valDuplicate: 'Duplicate timestamps detected',
      valDays: 'Days found',
      valComplete: 'Complete days in sizing window',
      valIncomplete: 'Incomplete days excluded',
      valMissing: 'Missing window intervals (across incomplete days)',
      valInterval: 'Interval (minutes) must be a whole number between 1 and 480 that divides 480 (e.g. 30).',
      valTarget: 'Target grid load must be a number of 0 or more.',
      valReserve: 'Reserve margin must be between 0 and 100%.',
      valConstKw: 'Discharge power must be greater than 0 kW.',
      mappingTsAnchor: 'Interval timestamp',
      mappingTsAnchorEnd: 'End time',
      mappingTsAnchorStart: 'Start time'
    }
  };

  var LANG = 'en';
  function t(k, vars) {
    var s = STR.en[k] || k;
    if (vars) { for (var v in vars) s = s.split('{' + v + '}').join(String(vars[v])); }
    return s;
  }

  /* ------------------------- helpers ------------------------- */
  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function num(v, dflt) { var n = parseFloat(v); return isFinite(n) ? n : dflt; }
  function optNum(v, factor) { var n = num(v, NaN); return isFinite(n) ? n * factor : null; }
  function fmt(n, d) {
    if (n == null || !isFinite(n)) return '\u2014';
    return Number(n).toLocaleString('en-US', { maximumFractionDigits: (d == null ? 1 : d) });
  }
  function fmtDate(d) {
    if (!d) return '\u2014';
    return d.toLocaleDateString('en-GB', { year: 'numeric', month: 'short', day: 'numeric' });
  }
  function dateFromSerial(serial) {
    var ms = Math.round((serial - 25569) * 86400000);
    var d = new Date(ms);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds());
  }
  function parseDateTimeString(s) {
    s = String(s).trim();
    var m = /^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(s);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})[ T]?(\d{1,2})?:?(\d{2})?/.exec(s);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), 0);
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  function parseTs(raw) {
    if (raw == null || raw === '') return null;
    if (typeof raw === 'number') return dateFromSerial(raw);
    var s = String(raw).trim();
    if (/^\d+(\.\d+)?$/.test(s)) return dateFromSerial(parseFloat(s));
    return parseDateTimeString(s);
  }
  function combineDateTime(d, t) {
    var base = null;
    if (d != null && d !== '') {
      if (typeof d === 'number') base = dateFromSerial(d); else base = parseDateTimeString(d);
      if (base) base = new Date(base.getFullYear(), base.getMonth(), base.getDate());
    }
    var h = 0, mi = 0, sec = 0;
    if (t != null && t !== '') {
      if (typeof t === 'number') {
        var frac = (t % 1 + 1) % 1;
        h = Math.floor(frac * 24); mi = Math.floor((frac * 24 - h) * 60); sec = Math.round(((frac * 24 - h) * 60 - mi) * 60);
      } else {
        var m = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(String(t));
        if (m) { h = +m[1]; mi = +m[2]; sec = +(m[3] || 0); }
      }
    }
    if (!base) return null;
    return new Date(base.getFullYear(), base.getMonth(), base.getDate(), h, mi, sec);
  }

  /* ------------------------- library ------------------------- */
  var LS_MODELS = 'bessSizing.library.v2';
  var LS_ACTIVE = 'bessSizing.activeModel.v2';
  var LS_CTRL = 'bessSizing.controlWindow.v1';
  var DEFAULT_MODEL = {
    id: 'custom-default', brand: '', name: 'Custom Model', grossKwh: 1000,
    socMinPct: 5, socMaxPct: 100, rtePct: 90, auxPct: 98,
    packPowerKw: 500, pcsPowerKw: 500,
    firstDegPct: 0, annualDegPct: 0, rteDefinition: 'AC-to-AC system RTE',
    degNotes: '', srcNotes: 'Default configurable pack'
  };
  var MODEL_HUAWEI = {
    id: 'huawei-luna-2000', brand: 'Huawei', name: 'Luna 2000-215',
    grossKwh: 215, socMinPct: 5, socMaxPct: 100, rtePct: 90, auxPct: 98,
    packPowerKw: 100, pcsPowerKw: 100,
    firstDegPct: 2, annualDegPct: 1, rteDefinition: 'AC-to-AC system RTE',
    degNotes: 'Illustrative reference values - verify against the current manufacturer datasheet before use.',
    srcNotes: 'Huawei Luna 2000-215 commercial & industrial ESS (approximate nameplate figures).'
  };
  var MODEL_BYD = {
    id: 'byd-mc-cube', brand: 'BYD', name: 'MC Cube',
    grossKwh: 262, socMinPct: 5, socMaxPct: 100, rtePct: 90, auxPct: 98,
    packPowerKw: 130, pcsPowerKw: 130,
    firstDegPct: 2, annualDegPct: 1, rteDefinition: 'AC-to-AC system RTE',
    degNotes: 'Illustrative reference values - verify against the current manufacturer datasheet before use.',
    srcNotes: 'BYD MC Cube commercial & industrial ESS (approximate nameplate figures).'
  };
  var MODEL_CATL = {
    id: 'catl-enerone', brand: 'CATL', name: 'EnerOne',
    grossKwh: 372, socMinPct: 5, socMaxPct: 100, rtePct: 90, auxPct: 98,
    packPowerKw: 186, pcsPowerKw: 186,
    firstDegPct: 2, annualDegPct: 1, rteDefinition: 'AC-to-AC system RTE',
    degNotes: 'Illustrative reference values - verify against the current manufacturer datasheet before use.',
    srcNotes: 'CATL EnerOne liquid-cooled rack (C&I) - approximate nameplate figures.'
  };
  var MODEL_TESLA = {
    id: 'tesla-megapack-2xl', brand: 'Tesla', name: 'Megapack 2 XL',
    grossKwh: 3900, socMinPct: 5, socMaxPct: 100, rtePct: 92, auxPct: 99,
    packPowerKw: 1925, pcsPowerKw: 1925,
    firstDegPct: 2, annualDegPct: 0.5, rteDefinition: 'AC-to-AC system RTE',
    degNotes: 'Illustrative reference values - verify against the current manufacturer datasheet before use.',
    srcNotes: 'Tesla Megapack 2 XL (utility/C&I) - approximate nameplate figures.'
  };
  var BUILTIN_MODELS = [DEFAULT_MODEL, MODEL_HUAWEI, MODEL_BYD, MODEL_CATL, MODEL_TESLA];
  function makeId() { return 'm' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
  function sanitizeModel(x) {
    var m = {};
    for (var k in DEFAULT_MODEL) m[k] = DEFAULT_MODEL[k];
    for (var k2 in (x || {})) { if (k2 in m) m[k2] = x[k2]; }
    m.grossKwh = num(m.grossKwh, 2250);
    m.socMinPct = num(m.socMinPct, 5);
    m.socMaxPct = num(m.socMaxPct, 100);
    m.rtePct = num(m.rtePct, 85);
    m.auxPct = num(m.auxPct, 100);
    m.packPowerKw = num(m.packPowerKw, 1125);
    m.pcsPowerKw = num(m.pcsPowerKw, 1125);
    m.firstDegPct = num(m.firstDegPct, 0);
    m.annualDegPct = num(m.annualDegPct, 0);
    m.rteDefinition = String(m.rteDefinition || 'AC-to-AC system RTE');
    if (m.socMinPct >= m.socMaxPct) { m.socMinPct = 5; m.socMaxPct = 100; }
    m.name = String(m.name || 'Custom Model');
    m.brand = String(m.brand || '');
    m.id = String(m.id || makeId());
    return m;
  }
  function toFinite(v) {
    if (v == null || v === '') return NaN;
    var n = typeof v === 'number' ? v : parseFloat(v);
    return isFinite(n) ? n : NaN;
  }
  function validateModel(x) {
    x = x || {};
    var errs = [];
    var gross = toFinite(x.grossKwh);
    var socMin = toFinite(x.socMinPct);
    var socMax = toFinite(x.socMaxPct);
    var rte = toFinite(x.rtePct);
    var aux = toFinite(x.auxPct);
    var pack = toFinite(x.packPowerKw);
    var pcs = toFinite(x.pcsPowerKw);
    var firstDeg = toFinite(x.firstDegPct == null ? 0 : x.firstDegPct);
    var annualDeg = toFinite(x.annualDegPct == null ? 0 : x.annualDegPct);
    if (!(gross > 0)) errs.push('valModelGross');
    if (!(socMin >= 0 && socMin <= 100)) errs.push('valModelSocMin');
    if (!(socMax >= 0 && socMax <= 100)) errs.push('valModelSocMax');
    if (isFinite(socMin) && isFinite(socMax) && socMin >= socMax) errs.push('valModelSocOrder');
    if (!(rte > 0 && rte <= 100)) errs.push('valModelRte');
    if (!(aux > 0 && aux <= 100)) errs.push('valModelAux');
    if (!(pack > 0)) errs.push('valModelPackPower');
    if (!(pcs > 0)) errs.push('valModelPcsPower');
    if (!(firstDeg >= 0 && firstDeg < 100)) errs.push('valModelFirstDeg');
    if (!(annualDeg >= 0 && annualDeg < 100)) errs.push('valModelAnnualDeg');
    return errs;
  }
  function modelDoD(m) { return Math.max(0, m.socMaxPct - m.socMinPct); }
  function modelUsable(m) { return m.grossKwh * modelDoD(m) / 100 * m.rtePct / 100 * m.auxPct / 100; }
  function modelSOH1(m) { return Math.max(0, 1 - m.firstDegPct / 100); }
  function modelSOH15(m) { return modelSOH1(m) * Math.pow(Math.max(0, 1 - m.annualDegPct / 100), 14); }
  function modelUsableYear1(m) { return modelUsable(m) * modelSOH1(m); }
  function modelUsableYear15(m) { return modelUsable(m) * modelSOH15(m); }

  function loadModels() {
    try {
      var raw = localStorage.getItem(LS_MODELS);
      if (raw) {
        var arr = JSON.parse(raw);
        if (Array.isArray(arr) && arr.length) {
          var out = [];
          var seenIds = {};
          for (var li = 0; li < arr.length; li++) {
            if (validateModel(arr[li]).length) continue;
            var lm = sanitizeModel(arr[li]);
            if (seenIds[lm.id]) continue;
            seenIds[lm.id] = true;
            out.push(lm);
          }
          if (out.length) return out;
        }
      }
    } catch (e) {}
    return BUILTIN_MODELS.map(sanitizeModel);
  }
  function saveModels() {
    try { localStorage.setItem(LS_MODELS, JSON.stringify(S.models)); } catch (e) {}
  }

  /* ------------------------- state ------------------------- */
  var S = {
    models: loadModels(),
    activeModelId: null,
    data: [],
    intervalHours: 0.5,
    unit: 'kW',
    strategy: 'static',
    target: 350,
    chartView: 'highest-md',
    sizingConvention: 'workbook',
    highestMD: null,
    reservePct: 10,
    constKw: 500,
    pvMode: 'none',
    pvMonthlyKwh: 0,
    pvPeakKw: 0,
    mapping: null,
    lastSheets: null,
    loadedSheets: [],
    lastSheetIndex: 0,
    validation: null,
    engine: null,
    sizing: null,
    winStart: '14:00',
    winEnd: '22:00',
    ctrlStart: '13:57',
    ctrlEnd: '22:03',
    controlWindow: null,
    tsAnchor: 'end'
  };
  try { S.activeModelId = localStorage.getItem(LS_ACTIVE); } catch (e) {}
  try {
    var _cw = JSON.parse(localStorage.getItem(LS_CTRL) || 'null');
    if (_cw && /^\d{1,2}:\d{2}$/.test(String(_cw.start || '')) && /^\d{1,2}:\d{2}$/.test(String(_cw.end || ''))) {
      S.ctrlStart = String(_cw.start);
      S.ctrlEnd = String(_cw.end);
    }
  } catch (e) {}

  function activeModel() {
    var m = null;
    for (var i = 0; i < S.models.length; i++) if (S.models[i].id === S.activeModelId) { m = S.models[i]; break; }
    if (!m) m = S.models[0];
    return m || sanitizeModel(DEFAULT_MODEL);
  }
  function setActiveModel(id) {
    S.activeModelId = id;
    try { localStorage.setItem(LS_ACTIVE, id); } catch (e) {}
    renderModelSelect();
    renderModelDetails();
    recompute();
  }
  /* ------------------------- i18n application ------------------------- */
  function applyLanguage() {
    document.documentElement.lang = 'en';
    var els = document.querySelectorAll('[data-i18n]');
    for (var i = 0; i < els.length; i++) {
      var el = els[i];
      var k = el.getAttribute('data-i18n');
      if (STR.en[k]) el.textContent = STR.en[k];
    }
    renderAll();
  }

  /* ------------------------- tabs ------------------------- */
  function initTabs() {
    var btns = document.querySelectorAll('.tab-btn');
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener('click', function () {
        var id = this.getAttribute('data-tab');
        var b2 = document.querySelectorAll('.tab-btn');
        for (var j = 0; j < b2.length; j++) b2[j].classList.remove('active');
        this.classList.add('active');
        var panes = document.querySelectorAll('.tab-pane');
        for (var k = 0; k < panes.length; k++) panes[k].classList.remove('active');
        $(id).classList.add('active');
        if (id === 'paneSizing') { renderChart(); renderDischargeSummary(); }
      });
    }
  }

  /* ------------------------- library UI ------------------------- */
  var editingId = null;
  function renderLibrary() {
    var tb = $('libraryTbody');
    var html = '';
    for (var i = 0; i < S.models.length; i++) {
      var m = S.models[i];
      var sel = m.id === S.activeModelId;
      var editing = m.id === editingId;
      var cls = (sel ? 'sel' : '') + (editing ? (sel ? ' edit' : 'edit') : '');
      html += '<tr class="' + cls + '">'
        + '<td>' + esc(m.name) + '</td>'
        + '<td>' + esc(m.brand || '') + '</td>'
        + '<td class="num">' + fmt(m.grossKwh, 0) + '</td>'
        + '<td class="num">' + fmt(modelDoD(m), 1) + '</td>'
        + '<td class="num">' + fmt(m.rtePct, 1) + '</td>'
        + '<td class="num">' + fmt(m.auxPct, 1) + '</td>'
        + '<td class="num">' + fmt(modelUsable(m), 0) + '</td>'
        + '<td class="num">' + fmt(modelSOH15(m) * 100, 2) + '%</td>'
        + '<td class="num">' + fmt(modelUsableYear15(m), 0) + '</td>'
        + '<td class="num">' + fmt(m.packPowerKw, 1) + '</td>'
        + '<td class="num">' + fmt(m.pcsPowerKw, 1) + '</td>'
        + '<td><button data-act="use" data-id="' + esc(m.id) + '">' + esc(t('libUse')) + '</button> '
        + '<button data-act="edit" data-id="' + esc(m.id) + '">' + esc(t('libEdit')) + '</button></td>'
        + '</tr>';
    }
    tb.innerHTML = html;
  }
  function fillForm(m) {
    editingId = m.id;
    $('mName').value = m.name;
    $('mBrand').value = m.brand || '';
    $('mGross').value = m.grossKwh;
    $('mSocMin').value = m.socMinPct;
    $('mSocMax').value = m.socMaxPct;
    $('mRte').value = m.rtePct;
    $('mAux').value = m.auxPct;
    $('mPackPower').value = m.packPowerKw;
    $('mPcsPower').value = m.pcsPowerKw;
    $('mFirstDeg').value = m.firstDegPct;
    $('mAnnualDeg').value = m.annualDegPct;
    $('mRteDefinition').value = m.rteDefinition || '';
    $('mDegNotes').value = m.degNotes || '';
    $('mSrcNotes').value = m.srcNotes || '';
    renderModelDegradationAudit(m);
  }
  function renderModelDegradationAudit(model) {
    var m = model || sanitizeModel({
      grossKwh: $('mGross').value, socMinPct: $('mSocMin').value, socMaxPct: $('mSocMax').value,
      rtePct: $('mRte').value, auxPct: $('mAux').value, packPowerKw: $('mPackPower').value,
      pcsPowerKw: $('mPcsPower').value, firstDegPct: $('mFirstDeg').value,
      annualDegPct: $('mAnnualDeg').value, rteDefinition: $('mRteDefinition').value
    });
    var el = $('modelDegradationAudit');
    if (!el) return;
    el.innerHTML = '<b>' + fmt(modelUsableYear15(m), 0) + ' kWh</b>'
      + 'Base usable: ' + fmt(modelUsable(m), 0) + ' kWh · Year 1: ' + fmt(modelUsableYear1(m), 0)
      + ' kWh · SOH Year 15: ' + fmt(modelSOH15(m) * 100, 2) + '%';
  }
  function readForm() {
    var raw = {
      id: editingId,
      name: $('mName').value,
      brand: $('mBrand').value,
      grossKwh: $('mGross').value,
      socMinPct: $('mSocMin').value,
      socMaxPct: $('mSocMax').value,
      rtePct: $('mRte').value,
      auxPct: $('mAux').value,
      packPowerKw: $('mPackPower').value,
      pcsPowerKw: $('mPcsPower').value,
      firstDegPct: $('mFirstDeg').value,
      annualDegPct: $('mAnnualDeg').value,
      rteDefinition: $('mRteDefinition').value,
      degNotes: $('mDegNotes').value,
      srcNotes: $('mSrcNotes').value
    };
    var errs = validateModel(raw);
    if (errs.length) {
      var msgs = [];
      for (var ei = 0; ei < errs.length; ei++) msgs.push(t(errs[ei]));
      status(msgs.join(' | '));
      return null;
    }
    return sanitizeModel(raw);
  }
  function status(msg) { $('libraryStatus').textContent = msg || ''; }
  var LIBRARY_XLSX_HEADERS = [
    'Model_ID','Brand','Model_Name','Gross_Pack_kWh','Min_SOC_pct','Max_SOC_pct','RTE_pct',
    'Aux_Retention_pct','Pack_Rated_Power_kW','PCS_Rated_Power_kW',
    'First_Year_Degradation_pct','Annual_Degradation_pct_Years_2_to_15','RTE_Definition',
    'Degradation_Notes','Source_Notes'
  ];
  function downloadExcelLibraryTemplate() {
    var b64 = window.__BESS_LIBRARY_TEMPLATE_BASE64__ || '';
    if (!b64) { status('Excel template is not embedded.'); return; }
    var bin = atob(b64), bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = 'bess_library_template.xlsx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(a.href);
  }
  function importExcelLibrary(file) {
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = window.XLSX.read(new Uint8Array(e.target.result), { type: 'array', raw: true });
        var sn = wb.SheetNames.indexOf('Battery_PCS_Library') >= 0 ? 'Battery_PCS_Library' : wb.SheetNames[0];
        var aoa = window.XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' });
        var headerRow = -1, map = {};
        for (var r = 0; r < Math.min(aoa.length, 20); r++) {
          map = {}; for (var c = 0; c < (aoa[r] || []).length; c++) map[String(aoa[r][c]).trim()] = c;
          if (map.Model_ID != null && map.Model_Name != null) { headerRow = r; break; }
        }
        if (headerRow < 0) throw new Error(t('libExcelMissingHeaders', { headers: 'Model_ID, Model_Name' }));
        var missing = LIBRARY_XLSX_HEADERS.filter(function (h) { return map[h] == null; });
        if (missing.length) throw new Error(t('libExcelMissingHeaders', { headers: missing.join(', ') }));
        var models = [], seen = {}, rowErrors = [];
        for (var rr = headerRow + 1; rr < aoa.length; rr++) {
          var row = aoa[rr] || [], id = String(row[map.Model_ID] || '').trim();
          if (!id) continue;
          var raw = {
            id: id, brand: row[map.Brand], name: row[map.Model_Name], grossKwh: row[map.Gross_Pack_kWh],
            socMinPct: row[map.Min_SOC_pct], socMaxPct: row[map.Max_SOC_pct], rtePct: row[map.RTE_pct],
            auxPct: row[map.Aux_Retention_pct], packPowerKw: row[map.Pack_Rated_Power_kW],
            pcsPowerKw: row[map.PCS_Rated_Power_kW], firstDegPct: row[map.First_Year_Degradation_pct],
            annualDegPct: row[map.Annual_Degradation_pct_Years_2_to_15], rteDefinition: row[map.RTE_Definition],
            degNotes: row[map.Degradation_Notes], srcNotes: row[map.Source_Notes]
          };
          var errs = validateModel(raw);
          if (!String(raw.name || '').trim()) errs.push('Model_Name');
          if (seen[id]) errs.push('duplicate Model_ID');
          if (errs.length) { rowErrors.push('Row ' + (rr + 1) + ': ' + errs.join(', ')); continue; }
          seen[id] = true; models.push(sanitizeModel(raw));
        }
        if (rowErrors.length || !models.length) throw new Error(t('libExcelRowErrors', { n: rowErrors.length || 1, detail: rowErrors.slice(0, 3).join(' | ') }));
        S.models = models; S.activeModelId = models[0].id; editingId = null; saveModels();
        renderLibrary(); renderModelSelect(); renderModelDetails(); fillForm(models[0]); recompute();
        status(t('libExcelImported', { n: models.length }));
      } catch (err) { status(err && err.message ? err.message : String(err)); }
    };
    reader.readAsArrayBuffer(file);
  }
  function initLibrary() {
    $('modelNew').addEventListener('click', function () {
      editingId = null;
      fillForm(sanitizeModel({ id: makeId(), name: 'Custom Model ' + (S.models.length + 1) }));
      renderLibrary();
      status('');
    });
    $('modelSave').addEventListener('click', function () {
      var m = readForm();
      if (!m) return;
      var idx = -1;
      for (var i = 0; i < S.models.length; i++) if (S.models[i].id === editingId) { idx = i; break; }
      if (idx >= 0) S.models[idx] = m; else S.models.push(m);
      editingId = m.id;
      saveModels();
      renderLibrary();
      if (S.activeModelId !== m.id) setActiveModel(m.id); else { renderModelSelect(); renderModelDetails(); recompute(); }
      status(t('libSaved'));
    });
    $('modelDelete').addEventListener('click', function () {
      if (editingId == null) return;
      if (S.models.length <= 1) { status(t('libResetDone')); return; }
      S.models = S.models.filter(function (x) { return x.id !== editingId; });
      if (S.activeModelId === editingId) S.activeModelId = S.models[0].id;
      editingId = null;
      saveModels();
      renderLibrary();
      renderModelSelect();
      renderModelDetails();
      recompute();
      status('');
    });
    $('modelDuplicate').addEventListener('click', function () {
      if (editingId == null) return;
      var src = null;
      for (var i = 0; i < S.models.length; i++) if (S.models[i].id === editingId) { src = S.models[i]; break; }
      if (!src) return;
      var copy = sanitizeModel(src);
      copy.id = makeId();
      copy.name = src.name + ' (copy)';
      S.models.push(copy);
      editingId = copy.id;
      saveModels();
      renderLibrary();
      fillForm(copy);
      status('');
    });
    $('libraryTbody').addEventListener('click', function (e) {
      var btn = e.target && e.target.closest ? e.target.closest('button') : null;
      if (!btn) return;
      var id = btn.getAttribute('data-id');
      var act = btn.getAttribute('data-act');
      var m = null;
      for (var i = 0; i < S.models.length; i++) if (S.models[i].id === id) { m = S.models[i]; break; }
      if (!m) return;
      if (act === 'use') setActiveModel(id);
      else if (act === 'edit') {
        fillForm(m);
        renderLibrary();
        status(t('libEditing', { name: m.name }));
        $('modelForm').scrollIntoView({ behavior: 'smooth', block: 'start' });
        var nm = $('mName');
        if (nm) { nm.focus(); nm.select(); }
      }
    });
    ['mGross','mSocMin','mSocMax','mRte','mAux','mPackPower','mPcsPower','mFirstDeg','mAnnualDeg','mRteDefinition'].forEach(function (id) {
      $(id).addEventListener('input', function () { renderModelDegradationAudit(); });
    });
    $('libReset').addEventListener('click', function () {
      S.models = BUILTIN_MODELS.map(sanitizeModel);
      S.activeModelId = S.models[0].id;
      editingId = null;
      saveModels();
      renderLibrary(); renderModelSelect(); renderModelDetails(); recompute();
      status(t('libResetDone'));
    });
    $('libExcelDownload').addEventListener('click', downloadExcelLibraryTemplate);
    $('libExcelImport').addEventListener('click', function () { $('libExcelFile').click(); });
    $('libExcelFile').addEventListener('change', function (e) { importExcelLibrary(e.target.files[0]); e.target.value = ''; });
  }

  /* ------------------------- model select/details ------------------------- */
  function renderModelSelect() {
    var sel = $('modelSelect');
    var html = '';
    for (var i = 0; i < S.models.length; i++) {
      var m = S.models[i];
      html += '<option value="' + esc(m.id) + '"' + (m.id === S.activeModelId ? ' selected' : '') + '>' + esc(m.name) + '</option>';
    }
    sel.innerHTML = html;
  }
  function renderModelDetails() {
    var m = activeModel();
    var rows = [
      ['detailUsable', fmt(modelUsable(m), 0) + ' kWh', true],
      ['detailUsableYear1', fmt(modelUsableYear1(m), 0) + ' kWh', false],
      ['detailUsableYear15', fmt(modelUsableYear15(m), 0) + ' kWh', true],
      ['detailSOH15', fmt(modelSOH15(m) * 100, 2) + '%', false],
      ['detailGross', fmt(m.grossKwh, 0) + ' kWh', false],
      ['detailDoD', fmt(modelDoD(m), 1) + '%', false],
      ['detailRte', fmt(m.rtePct, 1) + '%', false],
      ['detailAux', fmt(m.auxPct, 1) + '%', false],
      ['detailPackPower', fmt(m.packPowerKw, 1) + ' kW', false],
      ['detailPcsPower', fmt(m.pcsPowerKw, 1) + ' kW', false]
    ];
    var html = '';
    for (var i = 0; i < rows.length; i++) {
      html += '<div><span>' + esc(t(rows[i][0])) + '</span><b class="' + (rows[i][2] ? 'hi' : '') + '">' + esc(rows[i][1]) + '</b></div>';
    }
    $('modelDetails').innerHTML = html;
  }

  /* ------------------------- CSV template ------------------------- */
  function initCsvTemplate() {
    $('csvTemplateBtn').addEventListener('click', function () {
      var rows = [
        'timestamp,load_kw,pv_gen_kw,grid_export_kw',
        '2026-05-01 00:00,120,,',
        '2026-05-01 00:30,115,,',
        '2026-05-01 14:00,520,,',
        '2026-05-01 14:30,540,,'
      ];
      var blob = new Blob([rows.join('\n')], { type: 'text/csv' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'bess_load_template.csv';
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(a.href);
    });
  }
  /* ------------------------- header classification ------------------------- */
  function normHeader(v) { return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]+/g, ' '); }
  function hasToken(h, toks) { for (var i = 0; i < toks.length; i++) if (h.indexOf(toks[i]) !== -1) return true; return false; }
  var COMBINED_TOKENS = ['timestamp', 'datetime', 'date time', 'date end', 'end time', 'start time', 'date and time', 'time stamp'];
  function detectAnchor(headerCell) {
    var h = normHeader(headerCell);
    if (!h) return null;
    var hasStart = /\bstart\b/.test(h);
    var hasEnd = /\bend\b/.test(h);
    if (hasStart && !hasEnd) return 'start';
    if (hasEnd && !hasStart) return 'end';
    return null;
  }
  function detectAnchorFromHeader(row, c) {
    if (c.ts >= 0) return detectAnchor(row[c.ts]);
    if (c.time >= 0) return detectAnchor(row[c.time]);
    return null;
  }
  function classifyHeader(row) {
    var out = { ts: -1, date: -1, time: -1, load: -1, loadScore: 0, pv: -1, export: -1 };
    for (var i = 0; i < row.length; i++) {
      var h = normHeader(row[i]);
      if (!h) continue;
      if (hasToken(h, COMBINED_TOKENS) || (hasToken(h, ['date']) && hasToken(h, ['time']))) { if (out.ts < 0) out.ts = i; continue; }
      if (hasToken(h, ['date']) && out.date < 0) { out.date = i; continue; }
      if (hasToken(h, ['time']) && out.time < 0) { out.time = i; continue; }
      var score = 0;
      if (hasToken(h, ['kw import'])) score = 100;
      else if (hasToken(h, ['import']) && !hasToken(h, ['kvar', 'kvarh', 'export'])) score = 60;
      else if (hasToken(h, ['load'])) score = 70;
      else if (hasToken(h, ['demand'])) score = 70;
      else if (hasToken(h, ['kw']) && !hasToken(h, ['kvar', 'kvarh', 'export'])) score = 40;
      else if (hasToken(h, ['power']) && !hasToken(h, ['reactive', 'export'])) score = 30;
      if (score > out.loadScore) { out.loadScore = score; out.load = i; }
      if (hasToken(h, ['pv', 'solar', 'generation']) && out.pv < 0) out.pv = i;
      if (hasToken(h, ['export']) && out.export < 0) out.export = i;
    }
    return out;
  }
  function detectHeader(aoa) {
    var best = null;
    var limit = Math.min(aoa.length, 20);
    for (var r = 0; r < limit; r++) {
      var c = classifyHeader(aoa[r]);
      if (c.load < 0) continue;
      var hasTs = c.ts >= 0 || (c.date >= 0 && c.time >= 0);
      if (!hasTs) continue;
      var anchor = detectAnchorFromHeader(aoa[r], c);
      var score = c.loadScore + (c.ts >= 0 ? 100 : 80);
      if (anchor == null) score = Math.min(score, 119);
      if (!best || score > best.score) best = { row: r, c: c, score: score, anchor: anchor };
    }
    return best;
  }
  function inferUnit(headerCell) {
    var h = normHeader(headerCell);
    if (h.indexOf('mw') !== -1) return 'MW';
    if (h.indexOf('kva') !== -1 || h.indexOf('kvar') !== -1) return 'kW';
    return 'kW';
  }
  function inferIntervalMinutes(rows) {
    var diffs = [];
    for (var i = 1; i < Math.min(rows.length, 800); i++) {
      var d = (rows[i].ts - rows[i - 1].ts) / 60000;
      if (d > 0 && d <= 1440) diffs.push(d);
    }
    if (!diffs.length) return 30;
    diffs.sort(function (a, b) { return a - b; });
    var med = diffs[Math.floor(diffs.length / 2)];
    var snaps = [1, 5, 10, 15, 30, 60, 120, 1440];
    var best = snaps[0], bd = 1e9;
    for (var s = 0; s < snaps.length; s++) { var dd = Math.abs(med - snaps[s]); if (dd < bd) { bd = dd; best = snaps[s]; } }
    return best;
  }

  /* ------------------------- upload ------------------------- */
  function populateMapping(aoa) {
    var det = detectHeader(aoa);
    var headerRow = det ? det.row : 0;
    var c = det ? det.c : classifyHeader(aoa[0] || []);
    var sh = $('mapSheet');
    sh.innerHTML = '';
    for (var i = 0; i < S.mapping.sheets.length; i++) {
      var nm = S.mapping.sheets[i].name;
      sh.innerHTML += '<option value="' + i + '">' + esc(nm) + '</option>';
    }
    sh.value = String(S.mapping.sheetIndex);
    $('mapHeaderRow').value = headerRow + 1;
    setTsMode(c.ts >= 0 ? 'combined' : 'split');
    fillColumnSelect('mapTsCol', aoa[headerRow] || [], c.ts >= 0 ? c.ts : c.date);
    fillColumnSelect('mapDateCol', aoa[headerRow] || [], c.date);
    fillColumnSelect('mapTimeCol', aoa[headerRow] || [], c.time);
    fillColumnSelect('mapLoadCol', aoa[headerRow] || [], c.load);
    fillColumnSelect('mapPvCol', aoa[headerRow] || [], c.pv);
    fillColumnSelect('mapExportCol', aoa[headerRow] || [], c.export);
    $('mapUnit').value = inferUnit((aoa[headerRow] || [])[c.load]);
    var detAnchor = detectAnchorFromHeader(aoa[headerRow] || [], c);
    $('mapTsAnchor').value = detAnchor === 'start' ? 'start' : 'end';
    if (detAnchor == null) { $('mapConfidence').style.display = 'block'; $('mapConfidence').textContent = t('upLowConfidence'); }
    else { $('mapConfidence').style.display = 'none'; $('mapConfidence').textContent = ''; }
    $('mapInterval').value = 30;
    var probe = parseWithMapping(buildMappingFromForm());
    if (probe.length >= 2) $('mapInterval').value = inferIntervalMinutes(probe);
    $('mappingPanel').style.display = 'block';
  }
  function setTsMode(mode) {
    $('mapTsMode').value = mode;
    $('mapTsColWrap').style.display = mode === 'combined' ? 'block' : 'none';
    $('mapDateColWrap').style.display = mode === 'split' ? 'block' : 'none';
    $('mapTimeColWrap').style.display = mode === 'split' ? 'block' : 'none';
  }
  function fillColumnSelect(id, headerRow, selectedIdx) {
    var sel = $(id);
    var html = '<option value="-1">' + esc(t('mappingNone')) + '</option>';
    for (var i = 0; i < headerRow.length; i++) {
      var label = String(headerRow[i] == null ? '' : headerRow[i]).trim() || ('Column ' + (i + 1));
      html += '<option value="' + i + '"' + (i === selectedIdx ? ' selected' : '') + '>' + esc(label) + '</option>';
    }
    sel.innerHTML = html;
  }
  function buildMappingFromForm() {
    var aoa = S.mapping.sheets[parseInt($('mapSheet').value, 10)].aoa;
    var mode = $('mapTsMode').value;
    return {
      aoa: aoa,
      headerRow: (parseInt($('mapHeaderRow').value, 10) || 1) - 1,
      tsMode: mode,
      tsCol: parseInt($('mapTsCol').value, 10),
      dateCol: parseInt($('mapDateCol').value, 10),
      timeCol: parseInt($('mapTimeCol').value, 10),
      loadCol: parseInt($('mapLoadCol').value, 10),
      pvCol: parseInt($('mapPvCol').value, 10),
      exportCol: parseInt($('mapExportCol').value, 10),
      unit: $('mapUnit').value,
      tsAnchor: $('mapTsAnchor').value === 'start' ? 'start' : 'end',
      intervalMinutes: parseFloat($('mapInterval').value)
    };
  }
  function parseWithMapping(mapping) {
    var unitFactor = { kW: 1, MW: 1000, W: 0.001 }[mapping.unit] || 1;
    var rows = mapping.aoa.slice(mapping.headerRow + 1);
    var out = [];
    var seen = {};
    var stats = { raw: rows.length, parsed: 0, invalid: 0, duplicate: 0 };
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var ts = null;
      if (mapping.tsMode === 'combined') ts = parseTs(r[mapping.tsCol]);
      else ts = combineDateTime(r[mapping.dateCol], r[mapping.timeCol]);
      if (!ts) continue;
      var load = num(r[mapping.loadCol], NaN);
      if (!isFinite(load) || load < 0) { stats.invalid++; continue; }
      load = load * unitFactor;
      var key = ts.getTime();
      if (Object.prototype.hasOwnProperty.call(seen, key)) stats.duplicate++;
      else seen[key] = true;
      out.push({
        ts: ts,
        load: load,
        pv: optNum(r[mapping.pvCol], unitFactor),
        export: optNum(r[mapping.exportCol], unitFactor)
      });
    }
    out.sort(function (a, b) { return a.ts - b.ts; });
    stats.parsed = out.length;
    out.stats = stats;
    return out;
  }
  function applyDetectedSheets(sheets) {
    var merged = [], stats = { raw: 0, parsed: 0, invalid: 0, duplicate: 0 }, base = null, usedSheets = [];
    for (var si = 0; si < sheets.length; si++) {
      var det = detectHeader(sheets[si].aoa);
      if (!det || det.score < 120) continue;
      var c = det.c, headerRow = det.row;
      var mapping = {
        aoa: sheets[si].aoa, headerRow: headerRow, tsMode: c.ts >= 0 ? 'combined' : 'split',
        tsCol: c.ts, dateCol: c.date, timeCol: c.time, loadCol: c.load, pvCol: c.pv, exportCol: c.export,
        unit: inferUnit((sheets[si].aoa[headerRow] || [])[c.load]), tsAnchor: det.anchor || 'end', intervalMinutes: 30
      };
      var rows = parseWithMapping(mapping);
      if (!rows.length) continue;
      mapping.intervalMinutes = inferIntervalMinutes(rows);
      if (validateInterval(mapping.intervalMinutes) == null) continue;
      if (!base) base = mapping;
      if (mapping.intervalMinutes !== base.intervalMinutes || mapping.tsAnchor !== base.tsAnchor || mapping.unit !== base.unit) continue;
      for (var ri = 0; ri < rows.length; ri++) { rows[ri].sourceSheet = sheets[si].name; merged.push(rows[ri]); }
      stats.raw += rows.stats.raw; stats.parsed += rows.stats.parsed; stats.invalid += rows.stats.invalid; stats.duplicate += rows.stats.duplicate;
      usedSheets.push(sheets[si].name);
    }
    if (!merged.length || !base) return false;
    merged.sort(function (a, b) { return a.ts - b.ts; }); merged.stats = stats;
    S.loadedSheets = usedSheets;
    commitData(merged, base.intervalMinutes, base.unit, base.tsAnchor);
    return true;
  }
  function handleFile(file) {
    if (!file) return;
    $('uploadStatus').className = 'import-status';
    $('uploadStatus').textContent = t('upLoading');
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var wb = (window.XLSX && window.XLSX.read) ? window.XLSX.read(data, { type: 'array', raw: true }) : null;
        if (!wb || !wb.SheetNames || !wb.SheetNames.length) { fail('upNoData'); return; }
        var sheets = [];
        for (var i = 0; i < wb.SheetNames.length; i++) {
          var ws = wb.Sheets[wb.SheetNames[i]];
          var aoa = (window.XLSX.utils && window.XLSX.utils.sheet_to_json) ? window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }) : [];
          sheets.push({ name: wb.SheetNames[i], aoa: aoa });
        }
        S.lastSheets = sheets;
        var best = null;
        for (var s = 0; s < sheets.length; s++) {
          var det = detectHeader(sheets[s].aoa);
          if (det && (!best || det.score > best.score)) best = { sheetIndex: s, det: det };
        }
        if (!best) { fail('upNoData'); return; }
        S.lastSheetIndex = best.sheetIndex;
        var det = best.det;
        var headerRow = det.row;
        var c = det.c;
        var probeMapping = {
          aoa: sheets[best.sheetIndex].aoa,
          headerRow: headerRow,
          tsMode: c.ts >= 0 ? 'combined' : 'split',
          tsCol: c.ts, dateCol: c.date, timeCol: c.time,
          loadCol: c.load, pvCol: c.pv, exportCol: c.export,
          unit: inferUnit((sheets[best.sheetIndex].aoa[headerRow] || [])[c.load]),
          tsAnchor: det.anchor || 'end',
          intervalMinutes: 30
        };
        var parsed = parseWithMapping(probeMapping);
        if (parsed.length) probeMapping.intervalMinutes = inferIntervalMinutes(parsed);
        if (det.score >= 120) {
          if (!applyDetectedSheets(sheets)) applyDetected(sheets[best.sheetIndex].aoa, probeMapping);
        } else {
          S.mapping = { sheets: sheets, sheetIndex: best.sheetIndex };
          populateMapping(sheets[best.sheetIndex].aoa);
          $('uploadStatus').className = 'import-status warn';
          $('uploadStatus').textContent = t('upLowConfidence');
        }
      } catch (err) { fail('upError', { err: err && err.message ? err.message : String(err) }); }
    };
    reader.onerror = function () { fail('upError', { err: 'read error' }); };
    reader.readAsArrayBuffer(file);
  }
  function applyDetected(aoa, mapping) {
    S.mapping = null;
    $('mappingPanel').style.display = 'none';
    var parsed = parseWithMapping(mapping);
    if (!parsed.length) { fail('upNoData'); return; }
    commitData(parsed, mapping.intervalMinutes, mapping.unit, mapping.tsAnchor);
  }
  function applyMappingFromForm() {
    var mapping = buildMappingFromForm();
    var iv = validateInterval(mapping.intervalMinutes);
    if (iv == null) { fail('valInterval'); return; }
    mapping.intervalMinutes = iv;
    var parsed = parseWithMapping(mapping);
    if (!parsed.length) { fail('upNoData'); return; }
    S.mapping = null;
    $('mappingPanel').style.display = 'none';
    commitData(parsed, iv, mapping.unit, mapping.tsAnchor);
  }
  function commitData(rows, intervalMinutes, unit, anchor) {
    var iv = validateInterval(intervalMinutes);
    if (iv == null) { fail('valInterval'); return; }
    S.data = rows;
    S.intervalHours = iv / 60;
    S.unit = unit || 'kW';
    S.tsAnchor = anchor === 'start' ? 'start' : 'end';
    S.validation = buildValidation(rows, iv, S.tsAnchor);
    renderValidation();
    updatePvOptions();
    recompute();
    var days = S.validation ? S.validation.days : 0;
    $('uploadStatus').className = 'import-status ok';
    $('uploadStatus').textContent = t('upMapped', { n: rows.length, unit: S.unit, interval: iv, days: days });
  }
  function fail(key, vars) {
    $('uploadStatus').className = 'import-status err';
    $('uploadStatus').textContent = t(key, vars || {});
  }
  function initUpload() {
    var _fileInput = $('fileInput');
    if (window.bessAPI && window.bessAPI.openFile) {
      // Electron: use the native dialog via IPC (reliable for repeated uploads).
      _fileInput.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        window.bessAPI.openFile([{ name: 'Load profiles', extensions: ['xlsx', 'xls', 'csv'] }]).then(function (res) {
          if (!res || !res.base64) return;
          var bin = atob(res.base64);
          var bytes = new Uint8Array(bin.length);
          for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          handleFile(new File([bytes], res.name || 'profile.xlsx'));
        });
      });
    } else {
      // Browser: native file input.
      _fileInput.addEventListener('change', function (e) {
        if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
        setTimeout(function () { e.target.value = ''; }, 0);
      });
    }
    $('mapApply').addEventListener('click', applyMappingFromForm);
    $('mapCancel').addEventListener('click', function () { S.mapping = null; $('mappingPanel').style.display = 'none'; });
    $('mapAdjustBtn').addEventListener('click', function () {
      if (S.lastSheets && S.lastSheets.length) {
        S.mapping = { sheets: S.lastSheets, sheetIndex: S.lastSheetIndex || 0 };
        populateMapping(S.lastSheets[S.lastSheetIndex || 0].aoa);
      }
    });
    $('mapSheet').addEventListener('change', function () {
      var idx = parseInt($('mapSheet').value, 10);
      var aoa = S.mapping.sheets[idx].aoa;
      var det = detectHeader(aoa);
      var headerRow = det ? det.row : 0;
      var c = det ? det.c : classifyHeader(aoa[0] || []);
      $('mapHeaderRow').value = headerRow + 1;
      setTsMode(c.ts >= 0 ? 'combined' : 'split');
      fillColumnSelect('mapTsCol', aoa[headerRow] || [], c.ts >= 0 ? c.ts : c.date);
      fillColumnSelect('mapDateCol', aoa[headerRow] || [], c.date);
      fillColumnSelect('mapTimeCol', aoa[headerRow] || [], c.time);
      fillColumnSelect('mapLoadCol', aoa[headerRow] || [], c.load);
      fillColumnSelect('mapPvCol', aoa[headerRow] || [], c.pv);
      fillColumnSelect('mapExportCol', aoa[headerRow] || [], c.export);
      $('mapUnit').value = inferUnit((aoa[headerRow] || [])[c.load]);
      var detAnchor = detectAnchorFromHeader(aoa[headerRow] || [], c);
      $('mapTsAnchor').value = detAnchor === 'start' ? 'start' : 'end';
      if (detAnchor == null) { $('mapConfidence').style.display = 'block'; $('mapConfidence').textContent = t('upLowConfidence'); }
      else { $('mapConfidence').style.display = 'none'; $('mapConfidence').textContent = ''; }
    });
    $('mapTsMode').addEventListener('change', function () { setTsMode($('mapTsMode').value); });
  }
  /* ------------------------- validation ------------------------- */
  var WIN_START = 14 * 60, WIN_END = 22 * 60;

  function fmtHm(mins) {
    var m = Math.round(mins);
    return String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0');
  }
  function applySizingWindow() {
    var sm = parseHm(S.winStart), em = parseHm(S.winEnd);
    if (isFinite(sm) && isFinite(em) && em > sm) { WIN_START = sm; WIN_END = em; }
    var el = $('sizingWindowInfo');
    if (el) el.textContent = fmtHm(WIN_START) + ' - ' + fmtHm(WIN_END);
  }
  function onSizingWindowChanged() {
    applySizingWindow();
    if (S.data && S.data.length) {
      S.validation = buildValidation(S.data, S.intervalHours * 60, S.tsAnchor);
      renderValidation();
    }
    recompute();
  }

  function dayKeyOf(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }

  function validateInterval(minutes) {
    var v = toFinite(minutes);
    var wl = WIN_END - WIN_START;
    if (!isFinite(v) || v <= 0 || v !== Math.floor(v) || v > wl || (wl % v) !== 0) return null;
    return v;
  }
  function intervalBounds(ts, intervalMinutes, anchor) {
    var mins = ts.getHours() * 60 + ts.getMinutes() + ts.getSeconds() / 60;
    if (anchor === 'start') return { start: mins, end: mins + intervalMinutes };
    return { start: mins - intervalMinutes, end: mins };
  }
  function slotAnalysis(rows, intervalMinutes, anchor) {
    var iv = validateInterval(intervalMinutes);
    if (iv == null) {
      return { intervalMinutes: intervalMinutes, expectedCount: 0, expectedSlots: [], totalWindowDays: 0, completeDays: [], incompleteDays: [], rejected: [], days: {}, invalidInterval: true };
    }
    anchor = anchor === 'start' ? 'start' : 'end';
    var workbookMode = S.sizingConvention === 'workbook';
    var expectedSlots = [];
    if (workbookMode) for (var s = WIN_START; s <= WIN_END; s += iv) expectedSlots.push(s);
    else for (var s2 = WIN_START; s2 <= WIN_END - iv; s2 += iv) expectedSlots.push(s2);
    var expectedSet = {};
    for (var x = 0; x < expectedSlots.length; x++) expectedSet[expectedSlots[x]] = true;
    var days = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var key = dayKeyOf(r.ts);
      if (!days[key]) days[key] = { counts: 0, slots: {}, offGrid: [] };
      days[key].counts++;
      if (workbookMode) {
        var readingMinute = r.ts.getHours() * 60 + r.ts.getMinutes() + r.ts.getSeconds() / 60;
        if (readingMinute < WIN_START || readingMinute > WIN_END) continue;
        if (Object.prototype.hasOwnProperty.call(expectedSet, readingMinute)) days[key].slots[readingMinute] = (days[key].slots[readingMinute] || 0) + 1;
        else days[key].offGrid.push({ start: readingMinute, end: readingMinute, reason: 'offgrid-reading' });
        continue;
      }
      var b = intervalBounds(r.ts, iv, anchor);
      if (b.end <= WIN_START || b.start >= WIN_END) continue;
      if (b.start < WIN_START || b.end > WIN_END) {
        days[key].offGrid.push({ start: b.start, end: b.end, reason: 'boundary' });
        continue;
      }
      var sk = b.start;
      if (Object.prototype.hasOwnProperty.call(expectedSet, sk)) {
        days[key].slots[sk] = (days[key].slots[sk] || 0) + 1;
      } else {
        days[key].offGrid.push({ start: b.start, end: b.end, reason: 'offgrid' });
      }
    }
    var dayKeys = Object.keys(days).sort();
    var completeDays = [], incompleteDays = [], rejected = [];
    var perDay = {};
    for (var d = 0; d < dayKeys.length; d++) {
      var k = dayKeys[d];
      var missing = [], dup = [];
      for (var x2 = 0; x2 < expectedSlots.length; x2++) {
        var c = days[k].slots[expectedSlots[x2]] || 0;
        if (c === 0) missing.push(expectedSlots[x2]);
        else if (c > 1) dup.push(expectedSlots[x2]);
      }
      var offGrid = days[k].offGrid || [];
      var complete = missing.length === 0 && dup.length === 0 && offGrid.length === 0;
      var info = { counts: days[k].counts, missing: missing, dup: dup, offGrid: offGrid, complete: complete };
      perDay[k] = info;
      if (complete) completeDays.push(k);
      else incompleteDays.push(k);
    }
    for (var r3 = 0; r3 < incompleteDays.length; r3++) {
      var ik = incompleteDays[r3];
      rejected.push({ dayKey: ik, counts: perDay[ik].counts, missing: perDay[ik].missing, dup: perDay[ik].dup, offGrid: perDay[ik].offGrid });
    }
    return {
      intervalMinutes: iv,
      expectedCount: expectedSlots.length,
      expectedSlots: expectedSlots,
      totalWindowDays: dayKeys.length,
      completeDays: completeDays,
      incompleteDays: incompleteDays,
      rejected: rejected,
      days: perDay
    };
  }

  function buildValidation(rows, intervalMinutes, anchor) {
    var sa = slotAnalysis(rows, intervalMinutes, anchor);
    var missing = 0;
    for (var r2 = 0; r2 < sa.rejected.length; r2++) missing += sa.rejected[r2].missing.length;
    var totalDays = {};
    for (var j = 0; j < rows.length; j++) totalDays[dayKeyOf(rows[j].ts)] = 1;
    return {
      raw: rows.stats ? rows.stats.raw : rows.length,
      parsed: rows.length,
      invalid: rows.stats ? rows.stats.invalid : 0,
      duplicate: rows.stats ? rows.stats.duplicate : 0,
      days: Object.keys(totalDays).length,
      completeDays: sa.completeDays.length,
      incompleteDays: sa.incompleteDays.length,
      missingIntervals: missing,
      expected: sa.expectedCount,
      intervalMinutes: sa.intervalMinutes,
      slotEvidence: sa,
      completeDayKeys: sa.completeDays
    };
  }
  function renderValidation() {
    var v = S.validation;
    var el = $('validationSummary');
    if (!v) { el.style.display = 'none'; return; }
    var html = '<strong>' + esc(t('valTitle')) + '</strong><ul>'
      + '<li>' + esc(t('valRaw')) + ': ' + fmt(v.raw, 0) + '</li>'
      + '<li>' + esc(t('valParsed')) + ': ' + fmt(v.parsed, 0) + '</li>'
      + '<li class="' + (v.invalid ? 'warn' : 'ok') + '">' + esc(t('valInvalid')) + ': ' + fmt(v.invalid, 0) + '</li>'
      + '<li class="' + (v.duplicate ? 'warn' : 'ok') + '">' + esc(t('valDuplicate')) + ': ' + fmt(v.duplicate, 0) + '</li>'
      + '<li>' + esc(t('valDays')) + ': ' + fmt(v.days, 0) + '</li>'
      + '<li class="ok">' + esc(t('valComplete')) + ': ' + fmt(v.completeDays, 0) + '</li>'
      + '<li class="' + (v.incompleteDays ? 'warn' : 'ok') + '">' + esc(t('valIncomplete')) + ': ' + fmt(v.incompleteDays, 0) + '</li>'
      + (v.missingIntervals ? '<li class="warn">' + esc(t('valMissing')) + ': ' + fmt(v.missingIntervals, 0) + '</li>' : '')
      + '</ul>';
    el.innerHTML = html;
    el.style.display = 'block';
  }

  /* ------------------------- PV ------------------------- */
  function updatePvOptions() {
    var hasPv = false, hasExport = false;
    for (var i = 0; i < S.data.length; i++) {
      if (S.data[i].pv != null && isFinite(S.data[i].pv)) hasPv = true;
      if (S.data[i].export != null && isFinite(S.data[i].export)) hasExport = true;
    }
    var sel = $('pvModeSelect');
    var html = '<option value="none">' + esc(t('pvModeNone')) + '</option>';
    html += '<option value="detected"' + (hasPv ? '' : ' disabled') + '>' + esc(t('pvModeDetected')) + (hasPv ? '' : ' (\u2014)') + '</option>';
    html += '<option value="estimated">' + esc(t('pvModeEstimated')) + '</option>';
    sel.innerHTML = html;
    if (S.pvMode === 'detected' && !hasPv) S.pvMode = 'none';
    sel.value = S.pvMode;
  }
  function daysInMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate(); }
  function estimatePvKw(ts) {
    var monthly = S.pvMonthlyKwh || 0;
    var peak = S.pvPeakKw || 0;
    if (monthly <= 0) return 0;
    var minutes = ts.getHours() * 60 + ts.getMinutes();
    if (minutes < 6 * 60 || minutes > 18 * 60) return 0;
    var shape = Math.sin(Math.PI * (minutes - 6 * 60) / (12 * 60));
    var integralHours = 24 / Math.PI;
    var peakKw = peak > 0 ? peak : (monthly / daysInMonth(ts)) / integralHours;
    return shape * peakKw;
  }
  function effectiveLoad(r) {
    if (S.pvMode === 'detected' && r.pv != null && isFinite(r.pv)) {
      return r.load + r.pv - (isFinite(r.export) ? r.export : 0);
    }
    if (S.pvMode === 'estimated') {
      return r.load + estimatePvKw(r.ts) - (isFinite(r.export) ? r.export : 0);
    }
    return r.load;
  }

  /* ------------------------- engine ------------------------- */
  function computeEngine() {
    var intervalMinutes = S.intervalHours * 60;
    var anchor = S.tsAnchor === 'start' ? 'start' : 'end';
    var sa = S.validation && S.validation.slotEvidence ? S.validation.slotEvidence : slotAnalysis(S.data, intervalMinutes, anchor);
    var completeSet = {};
    for (var c = 0; c < sa.completeDays.length; c++) completeSet[sa.completeDays[c]] = true;
    var byDay = {};
    var used = {};
    var workbookMode = S.sizingConvention === 'workbook';
    var highestMD = null;
    for (var i = 0; i < S.data.length; i++) {
      var r = S.data[i];
      var key = dayKeyOf(r.ts);
      if (!completeSet[key]) continue;
      var loadValue = effectiveLoad(r);
      if (!highestMD || loadValue > highestMD.load) highestMD = { dayKey: key, ts: r.ts, load: loadValue };
      var slot;
      if (workbookMode) {
        slot = r.ts.getHours() * 60 + r.ts.getMinutes() + r.ts.getSeconds() / 60;
        if (slot < WIN_START || slot > WIN_END) continue;
      } else {
        var b = intervalBounds(r.ts, intervalMinutes, anchor);
        if (b.start < WIN_START || b.end > WIN_END) continue;
        slot = b.start;
      }
      if (!byDay[key]) { byDay[key] = []; used[key] = {}; }
      if (used[key][slot]) continue;
      used[key][slot] = true;
      byDay[key].push({ ts: r.ts, load: loadValue, slot: slot });
    }
    var gov = null;
    for (var dk in byDay) {
      var ints = byDay[dk];
      var target = S.target;
      var req = 0, maxNeed = 0;
      for (var j = 0; j < ints.length; j++) {
        var need = Math.max(ints[j].load - target, 0);
        req += need * S.intervalHours;
        if (need > maxNeed) maxNeed = need;
      }
      if (S.strategy === 'const') {
        req = S.constKw * (sa.expectedCount * S.intervalHours);
        maxNeed = S.constKw;
      }
      if (!gov || req > gov.req) gov = { dayKey: dk, ints: ints, req: req, maxNeed: maxNeed, target: target };
    }
    S.highestMD = highestMD;
    S.engine = { gov: gov, byDay: byDay, expected: sa.expectedCount, highestMD: highestMD };
    return S.engine;
  }
  function flatTopTarget(ints) {
    var loads = [];
    for (var i = 0; i < ints.length; i++) loads.push(ints[i].load);
    loads.sort(function (a, b) { return a - b; });
    if (!loads.length) return 0;
    var idx = Math.min(loads.length - 1, Math.floor(loads.length * 0.75));
    return loads[idx];
  }

  /* ------------------------- sizing ------------------------- */
  function controlErrors() {
    var errs = [];
    if (S.strategy !== 'const' && !(isFinite(S.target) && S.target >= 0)) errs.push('valTarget');
    if (!(isFinite(S.reservePct) && S.reservePct >= 0 && S.reservePct <= 100)) errs.push('valReserve');
    if (S.strategy === 'const' && !(isFinite(S.constKw) && S.constKw > 0)) errs.push('valConstKw');
    return errs;
  }
  function renderControlError(errs) {
    var el = $('controlError');
    if (!el) return;
    if (errs.length) {
      var msgs = [];
      for (var i = 0; i < errs.length; i++) msgs.push(t(errs[i]));
      el.style.display = 'block';
      el.textContent = msgs.join(' | ');
    } else {
      el.style.display = 'none';
      el.textContent = '';
    }
  }
  function computeSizing() {
    computeEngine();
    var m = activeModel();
    var gov = S.engine.gov;
    var errs = controlErrors();
    renderControlError(errs);
    if (errs.length) { S.sizing = null; return; }
    if (!gov) { S.sizing = null; return; }
    var usablePerPack = modelUsableYear15(m);
    var requiredEnergy = gov.req;
    var reserve = S.reservePct / 100;
    var energyPacks = Math.ceil(requiredEnergy * (1 + reserve) / usablePerPack);
    var nameplate = energyPacks * m.grossKwh;
    var usableBase = energyPacks * modelUsable(m);
    var usableYear15 = energyPacks * usablePerPack;
    var requiredPower = gov.maxNeed;
    var pcsQty = m.pcsPowerKw > 0 ? Math.ceil(requiredPower / m.pcsPowerKw) : 0;
    var powerPacks = m.packPowerKw > 0 ? Math.ceil(requiredPower / m.packPowerKw) : 0;
    var powerWarn = powerPacks > energyPacks;
    S.sizing = {
      gov: gov, requiredEnergy: requiredEnergy, energyPacks: energyPacks,
      nameplate: nameplate, usable: usableBase, usableYear15: usableYear15, usablePerPackYear15: usablePerPack, soh15: modelSOH15(m), requiredPower: requiredPower,
      pcsQty: pcsQty, powerPacks: powerPacks, powerWarn: powerWarn
    };
  }
  function recompute() {
    computeSizing();
    renderResults();
    renderChart();
    renderDischargeSummary();
    $('chartViewSelect').value = S.chartView;
    exposeDebug();
  }

  /* ------------------------- results ------------------------- */
  function renderResults() {
    var s = S.sizing;
    var grid = $('recoGrid');
    var noData = $('noData');
    if (!s) { grid.style.display = 'none'; noData.style.display = 'block'; $('resultWarn').style.display = 'none'; return; }
    noData.style.display = 'none';
    grid.style.display = 'grid';
    var day = s.gov.ints.length ? s.gov.ints[0].ts : null;
    $('recoGoverning').textContent = fmtDate(day);
    $('recoRequiredEnergy').textContent = fmt(s.requiredEnergy, 0) + ' kWh';
    $('recoEnergyPacks').textContent = fmt(s.energyPacks, 0);
    $('recoNameplate').textContent = fmt(s.nameplate, 0) + ' kWh';
    $('recoUsable').textContent = fmt(s.usable, 0) + ' kWh';
    $('recoYear15Usable').textContent = fmt(s.usableYear15, 0) + ' kWh';
    $('recoPower').textContent = fmt(s.requiredPower, 1) + ' kW';
    $('recoPcs').textContent = fmt(s.pcsQty, 0);
    $('recoPowerPacks').textContent = fmt(s.powerPacks, 0);
    if (s.powerWarn) {
      $('resultWarn').style.display = 'block';
      $('resultWarn').textContent = t('warnPower', { powerPacks: s.powerPacks, power: fmt(s.requiredPower, 1), energyPacks: s.energyPacks });
    } else {
      $('resultWarn').style.display = 'none';
    }
  }

  /* ------------------------- discharge verification (inline) ------------------------- */
  function dischargePowerLimit() {
    var m = activeModel();
    return S.sizing ? S.sizing.energyPacks * m.packPowerKw : 0;
  }
  function simDischarge(ints, capacity, target, powerLimit) {
    var energy = 0, dayBefore = 0, dayAfter = 0, disc = [], grid = [];
    for (var j = 0; j < ints.length; j++) {
      var load = ints[j].load;
      var d = Math.min(Math.max(load - target, 0), powerLimit);
      var remain = capacity - energy;
      if (remain <= 0) d = 0;
      else if (d * S.intervalHours > remain) d = remain / S.intervalHours;
      if (load > dayBefore) dayBefore = load;
      var g = load - d;
      if (g > dayAfter) dayAfter = g;
      disc.push(d); grid.push(g);
      energy += d * S.intervalHours;
    }
    return { disc: disc, grid: grid, dayBefore: dayBefore, dayAfter: dayAfter, energy: energy, remaining: capacity - energy };
  }
  function computeDischargeSummary() {
    if (!S.sizing || !S.engine || !S.engine.byDay) return null;
    var capacity = S.sizing.usable;
    var target = S.target;
    var powerLimit = dischargePowerLimit();
    var mdBefore = 0, mdAfter = 0, worstRemaining = capacity, worstDayKey = null;
    for (var dk in S.engine.byDay) {
      var ints = S.engine.byDay[dk].slice().sort(function (a, b) { return a.ts - b.ts; });
      var r = simDischarge(ints, capacity, target, powerLimit);
      if (r.dayBefore > mdBefore) mdBefore = r.dayBefore;
      if (r.dayAfter > mdAfter) mdAfter = r.dayAfter;
      if (r.remaining < worstRemaining) { worstRemaining = r.remaining; worstDayKey = dk; }
    }
    var reduction = mdBefore > 0 ? (mdBefore - mdAfter) / mdBefore * 100 : 0;
    var peakShaved = mdAfter <= target + 1e-6;
    return {
      mdBefore: mdBefore, mdAfter: mdAfter, reduction: reduction,
      sufficient: worstRemaining >= -1e-9 && peakShaved, worstDayKey: worstDayKey
    };
  }
  function renderDischargeSummary() {
    var el = $('mainDvSummary'), verdict = $('mainDvVerdict');
    if (!el) return;
    var dv = computeDischargeSummary();
    if (!dv) { el.style.display = 'none'; return; }
    el.style.display = 'flex';
    $('mainMdBefore').textContent = fmt(dv.mdBefore, 0) + ' kW';
    $('mainMdAfter').textContent = fmt(dv.mdAfter, 0) + ' kW';
    $('mainMdReduction').textContent = fmt(dv.reduction, 1) + '%';
    if (verdict) {
      verdict.style.display = 'inline-block';
      if (dv.sufficient) { verdict.className = 'dv-verdict ok'; verdict.textContent = t('dvSufficient'); }
      else { verdict.className = 'dv-verdict err'; verdict.textContent = t('dvInsufficient', { day: dv.worstDayKey }); }
    }
  }

  /* ------------------------- chart ------------------------- */
  var chartInstance = null;
  function round1(n) { return Math.round(n * 10) / 10; }
  function renderChart() {
    if (!window.Chart) { window.__BESS_CHART__ = null; return; }
    var cv = $('chart');
    var s = S.sizing;
    if ($('highestMdValue')) $('highestMdValue').textContent = S.highestMD ? (fmt(S.highestMD.load, 1) + ' kW · ' + fmtDate(S.highestMD.ts)) : '\u2014';
    if ($('governingEnergyValue')) $('governingEnergyValue').textContent = s && s.gov ? (fmtDate(s.gov.ints[0] && s.gov.ints[0].ts) + ' · ' + fmt(s.requiredEnergy, 0) + ' kWh') : '\u2014';
    if (!s || !s.gov) {
      if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
      window.__BESS_CHART__ = null;
      return;
    }
    var labels = [];
    var loads = [];
    var discharge = [];
    var gridAfter = [];
    var datasets;
    var isConst = S.strategy === 'const';
    var chartDayKey = (S.chartView === 'highest-md' && S.highestMD) ? S.highestMD.dayKey : s.gov.dayKey;
    if (isConst) {
      var intervalMinutes = S.intervalHours * 60;
      var anchor = S.tsAnchor === 'start' ? 'start' : 'end';
      var seen = {};
      for (var ci = 0; ci < S.data.length; ci++) {
        var cr = S.data[ci];
        if (dayKeyOf(cr.ts) !== chartDayKey) continue;
        var cb = intervalBounds(cr.ts, intervalMinutes, anchor);
        if (seen[cb.start]) continue;
        seen[cb.start] = true;
        labels.push(String(cr.ts.getHours()).padStart(2, '0') + ':' + String(cr.ts.getMinutes()).padStart(2, '0'));
        var cload = effectiveLoad(cr);
        var readingMinute = cr.ts.getHours() * 60 + cr.ts.getMinutes() + cr.ts.getSeconds() / 60;
        var inWindow = S.sizingConvention === 'workbook' ? (readingMinute >= WIN_START && readingMinute <= WIN_END) : (cb.start >= WIN_START && cb.end <= WIN_END);
        var battery = inWindow ? Math.min(cload, S.constKw) : 0;
        loads.push(round1(cload));
        discharge.push(round1(battery));
        gridAfter.push(round1(cload - battery));
      }
      datasets = [
        { label: t('legendLoad'), data: loads, borderColor: '#94a3b8', backgroundColor: 'rgba(148,163,184,.12)', pointRadius: 0, fill: true, tension: .25 },
        { label: t('legendBattery'), data: discharge, borderColor: '#10b981', pointRadius: 0, fill: false, tension: .25 },
        { label: t('legendGridAfter'), data: gridAfter, borderColor: '#3b82f6', pointRadius: 0, fill: false, tension: .25 }
      ];
    } else {
      var ints2 = (S.engine.byDay[chartDayKey] || s.gov.ints).slice().sort(function (a, b) { return a.ts - b.ts; });
      var sim = simDischarge(ints2, s.usable, S.target, dischargePowerLimit());
      for (var i = 0; i < ints2.length; i++) {
        var x = ints2[i];
        labels.push(String(x.ts.getHours()).padStart(2, '0') + ':' + String(x.ts.getMinutes()).padStart(2, '0'));
        loads.push(round1(x.load));
        discharge.push(round1(sim.disc[i]));
        gridAfter.push(round1(sim.grid[i]));
      }
      var target = S.target;
      datasets = [
        { label: t('legendLoad'), data: loads, borderColor: '#94a3b8', backgroundColor: 'rgba(148,163,184,.12)', pointRadius: 0, fill: true, tension: .25 },
        { label: t('legendTarget'), data: labels.map(function () { return target; }), borderColor: '#ef4444', borderDash: [6, 4], pointRadius: 0, fill: false },
        { label: t('legendBattery'), data: discharge, borderColor: '#10b981', pointRadius: 0, fill: false, tension: .25 },
        { label: t('legendGridAfter'), data: gridAfter, borderColor: '#3b82f6', pointRadius: 0, fill: false, tension: .25 }
      ];
    }
    if ($('legTarget')) $('legTarget').style.display = isConst ? 'none' : '';
    var ctx = cv.getContext('2d');
    if (chartInstance) chartInstance.destroy();
    chartInstance = new window.Chart(ctx, {
      type: 'line',
      data: { labels: labels, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { grid: { color: 'rgba(148,163,184,.30)' }, ticks: { color: '#5b6b80', maxTicksLimit: 12 } },
          y: { beginAtZero: true, grid: { color: 'rgba(148,163,184,.30)' }, ticks: { color: '#5b6b80' } }
        },
        plugins: { legend: { display: false } }
      }
    });
    window.__BESS_CHART__ = {
      rendered: true, version: (window.Chart.version || null), labelCount: labels.length, datasetCount: datasets.length,
      mode: isConst ? 'const' : S.strategy, chartDayKey: chartDayKey, labels: labels,
      datasets: datasets.map(function (d) { return { label: d.label, type: d.type || 'line', data: d.data }; })
    };
  }

  /* ------------------------- debug handle (test hook) ------------------------- */
  function exposeDebug() {
    var models = [];
    for (var i = 0; i < S.models.length; i++) {
      var mm = S.models[i];
      models.push({ id: mm.id, brand: mm.brand, name: mm.name, grossKwh: mm.grossKwh, socMinPct: mm.socMinPct, socMaxPct: mm.socMaxPct, rtePct: mm.rtePct, auxPct: mm.auxPct, packPowerKw: mm.packPowerKw, pcsPowerKw: mm.pcsPowerKw, firstDegPct: mm.firstDegPct, annualDegPct: mm.annualDegPct, rteDefinition: mm.rteDefinition, usable: modelUsable(mm), usableYear1: modelUsableYear1(mm), usableYear15: modelUsableYear15(mm), soh15: modelSOH15(mm), dod: modelDoD(mm) });
    }
    window.__BESS_APP__ = {
      lang: LANG,
      models: models,
      activeModelId: S.activeModelId,
      strategy: S.strategy, target: S.target, reservePct: S.reservePct, pvMode: S.pvMode, chartView: S.chartView, sizingConvention: S.sizingConvention, loadedSheets: S.loadedSheets, highestMD: S.highestMD,
      intervalHours: S.intervalHours, unit: S.unit, tsAnchor: S.tsAnchor,
      controlWindow: S.controlWindow,
      validation: S.validation,
      sizing: S.sizing ? {
        governingDay: fmtDate(S.sizing.gov.ints[0] && S.sizing.gov.ints[0].ts),
        governingDayKey: S.sizing.gov.dayKey,
        requiredEnergy: S.sizing.requiredEnergy,
        energyPacks: S.sizing.energyPacks,
        nameplate: S.sizing.nameplate,
        usable: S.sizing.usable,
        usableYear15: S.sizing.usableYear15,
        soh15: S.sizing.soh15,
        requiredPower: S.sizing.requiredPower,
        pcsQty: S.sizing.pcsQty,
        powerPacks: S.sizing.powerPacks,
        powerWarn: S.sizing.powerWarn
      } : null,
      rowCount: S.data.length
    };
  }

  /* ------------------------- controls ------------------------- */
  function parseHm(v) {
    var s = String(v == null ? '' : v).trim();
    var m = /^(\d{1,2}):(\d{2})$/.exec(s);
    if (!m) return NaN;
    var h = +m[1], mi = +m[2];
    if (h > 23 || mi > 59) return NaN;
    return h * 60 + mi;
  }
  function persistControlWindow() {
    try { localStorage.setItem(LS_CTRL, JSON.stringify({ start: S.ctrlStart, end: S.ctrlEnd })); } catch (e) {}
  }
  function updateControlWindow() {
    var sm = parseHm(S.ctrlStart), em = parseHm(S.ctrlEnd);
    var valid = isFinite(sm) && isFinite(em) && sm < em;
    S.controlWindow = { start: S.ctrlStart, end: S.ctrlEnd, valid: valid, startMin: sm, endMin: em, coversSizingWindow: valid && sm <= WIN_START && em >= WIN_END };
    var el = $('opReadiness');
    if (el) {
      if (!valid) { el.className = 'note err'; el.textContent = t('opReadinessErr'); }
      else if (S.controlWindow.coversSizingWindow) { el.className = 'note ok'; el.textContent = t('opReadinessOk'); }
      else { el.className = 'note warn'; el.textContent = t('opReadinessWarn'); }
    }
    exposeDebug();
  }
  function renderStrategyUi() {
    var st = S.strategy;
    $('strategySelect').value = st;
    $('targetInput').closest('div').style.display = st === 'const' ? 'none' : 'block';
    $('constKwWrap').style.display = st === 'const' ? 'block' : 'none';
    var noteKey = st === 'static' ? 'eduStatic' : (st === 'flattop' ? 'eduFlatTop' : 'eduConst');
    $('strategyNote').textContent = t(noteKey);
    $('autoTargetBtn').style.display = st === 'const' ? 'none' : 'inline-block';
  }
  function targetRequirement(target) {
    if (!S.engine || !S.engine.byDay) return null;
    var worstEnergy = 0, worstPower = 0;
    for (var dk in S.engine.byDay) {
      var ints = S.engine.byDay[dk], energy = 0, power = 0;
      for (var i = 0; i < ints.length; i++) {
        var need = Math.max(ints[i].load - target, 0);
        energy += need * S.intervalHours; if (need > power) power = need;
      }
      if (energy > worstEnergy) worstEnergy = energy;
      if (power > worstPower) worstPower = power;
    }
    return { energy: worstEnergy, power: worstPower };
  }
  function autoCalculateTarget() {
    recompute();
    if (!S.sizing || S.strategy === 'const') { $('autoTargetNote').textContent = t('autoTargetUnavailable'); return; }
    var n = S.sizing.energyPacks, m = activeModel(), reserve = 1 + S.reservePct / 100;
    var energyLimit = n * modelUsableYear15(m) / reserve;
    var powerLimit = n * m.packPowerKw;
    var maxLoad = S.highestMD ? Math.ceil(S.highestMD.load) : 0, found = null;
    for (var target = 0; target <= maxLoad; target++) {
      var req = targetRequirement(target);
      if (req && req.energy <= energyLimit + 1e-9 && req.power <= powerLimit + 1e-9) { found = target; break; }
    }
    if (found == null) { $('autoTargetNote').textContent = t('autoTargetUnavailable'); return; }
    S.target = found; $('targetInput').value = found; recompute();
    $('autoTargetNote').textContent = t('autoTargetUsing', { n: n, target: found });
  }
  function initControls() {
    $('strategySelect').addEventListener('change', function () { S.strategy = this.value; renderStrategyUi(); recompute(); });
    $('targetInput').addEventListener('input', function () { S.target = toFinite(this.value); recompute(); });
    $('constKwInput').addEventListener('input', function () { S.constKw = toFinite(this.value); recompute(); });
    $('reserveInput').addEventListener('input', function () { S.reservePct = toFinite(this.value); recompute(); });
    $('autoTargetBtn').addEventListener('click', autoCalculateTarget);
    $('chartViewSelect').addEventListener('change', function () { S.chartView = this.value; renderChart(); exposeDebug(); });
    $('modelSelect').addEventListener('change', function () { setActiveModel(this.value); });
    $('pvModeSelect').addEventListener('change', function () {
      S.pvMode = this.value;
      $('pvEstimatePanel').style.display = S.pvMode === 'estimated' ? 'block' : 'none';
      recompute();
    });
    $('pvMonthlyInput').addEventListener('input', function () { S.pvMonthlyKwh = num(this.value, 0); recompute(); });
    $('pvPeakKwInput').addEventListener('input', function () { S.pvPeakKw = num(this.value, 0); recompute(); });
    $('ctrlStartInput').addEventListener('input', function () { S.ctrlStart = this.value || '13:57'; persistControlWindow(); updateControlWindow(); });
    $('ctrlEndInput').addEventListener('input', function () { S.ctrlEnd = this.value || '22:03'; persistControlWindow(); updateControlWindow(); });
    if ($('winStartInput')) $('winStartInput').addEventListener('input', function () { S.winStart = this.value || '14:00'; onSizingWindowChanged(); });
    if ($('winEndInput')) $('winEndInput').addEventListener('input', function () { S.winEnd = this.value || '22:00'; onSizingWindowChanged(); });
    if ($('demandIntervalInput')) $('demandIntervalInput').addEventListener('input', function () {
      var mins = toFinite(this.value);
      if (isFinite(mins) && mins > 0) {
        S.intervalHours = mins / 60;
        if (S.data && S.data.length) { S.validation = buildValidation(S.data, mins, S.tsAnchor); renderValidation(); }
        recompute();
      }
    });
    $('applyReco').addEventListener('click', function () {
      recompute();
      if (S.sizing) {
        var el = $('uploadStatus');
        el.className = 'import-status ok';
        el.textContent = t('applied', { usable: fmt(S.sizing.usable, 0) });
      }
    });
  }

  /* ------------------------- init ------------------------- */
  function renderAll() {
    renderLibrary();
    renderModelSelect();
    renderModelDetails();
    renderStrategyUi();
    renderValidation();
    updatePvOptions();
    renderResults();
    renderChart();
    renderDischargeSummary();
    $('winStartInput').value = S.winStart;
    $('winEndInput').value = S.winEnd;
    $('demandIntervalInput').value = S.intervalHours * 60;
    applySizingWindow();
    $('ctrlStartInput').value = S.ctrlStart;
    $('ctrlEndInput').value = S.ctrlEnd;
    updateControlWindow();
    exposeDebug();
  }

  function boot() {
    var hasActive = false;
    for (var ai = 0; ai < S.models.length; ai++) if (S.models[ai].id === S.activeModelId) hasActive = true;
    if (!hasActive) S.activeModelId = S.models[0].id;
    initTabs();
    initLibrary();
    initUpload();
    initCsvTemplate();
    initControls();
    applyLanguage();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
