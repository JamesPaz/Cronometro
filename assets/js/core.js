/* =========================================================
   core.js - estado, reloj monotonico, posiciones y formato
   ========================================================= */
window.App = window.App || {};

(function (App) {
  'use strict';

  var Core = {};

  /* -------------------------------------------------- estado */
  Core.state = {
    roster: { teams: [] },   // equipos cargados
    current: null,          // prueba en curso
    history: [],            // pruebas finalizadas
    clock: { base: null, running: false, elapsedAtFinish: null },
    ui: { collapsed: [] }   // tarjetas del panel que estan colapsadas
  };

  /* -------------------------------------------------- utilidades */
  Core.pad = function (n, size) {
    var s = String(Math.abs(Math.floor(n)));
    while (s.length < size) { s = '0' + s; }
    return (n < 0 ? '-' : '') + s;
  };

  Core.uid = function () {
    return 'id-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
  };

  Core.escape = function (str) {
    return String(str == null ? '' : str)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  };

  Core.normalize = function (str) {
    return String(str == null ? '' : str)
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '');
  };

  /* -------------------------------------------------- formato de tiempo */
  // 00:00:00  (reloj principal)
  Core.fmtClock = function (ms) {
    ms = Math.max(0, Math.floor(ms || 0));
    var h = Math.floor(ms / 3600000);
    var m = Math.floor((ms % 3600000) / 60000);
    var s = Math.floor((ms % 60000) / 1000);
    return Core.pad(h, 2) + ':' + Core.pad(m, 2) + ':' + Core.pad(s, 2);
  };

  // .00  (centisegundos del reloj principal)
  Core.fmtCs = function (ms) {
    ms = Math.max(0, ms || 0);
    return '.' + Core.pad(Math.floor((ms % 1000) / 10), 2);
  };

  // 00:00:00.000  (tiempos exactos: inicio, fin, duracion)
  Core.fmtFull = function (ms) {
    if (ms == null || isNaN(ms)) { return ''; }
    ms = Math.max(0, ms);
    var h = Math.floor(ms / 3600000);
    var m = Math.floor((ms % 3600000) / 60000);
    var s = Math.floor((ms % 60000) / 1000);
    var cs = Math.floor(ms % 1000);
    return Core.pad(h, 2) + ':' + Core.pad(m, 2) + ':' + Core.pad(s, 2) + '.' + Core.pad(cs, 3);
  };

  Core.fmtDateTime = function (iso) {
    if (!iso) { return ''; }
    var d = new Date(iso);
    if (isNaN(d.getTime())) { return ''; }
    return Core.pad(d.getDate(), 2) + '/' + Core.pad(d.getMonth() + 1, 2) + '/' + d.getFullYear() +
      ' ' + Core.pad(d.getHours(), 2) + ':' + Core.pad(d.getMinutes(), 2);
  };

  Core.fmtFileStamp = function (iso) {
    var d = iso ? new Date(iso) : new Date();
    if (isNaN(d.getTime())) { d = new Date(); }
    return d.getFullYear() + Core.pad(d.getMonth() + 1, 2) + Core.pad(d.getDate(), 2) +
      '_' + Core.pad(d.getHours(), 2) + Core.pad(d.getMinutes(), 2);
  };

  /* -------------------------------------------------- reloj */
  Core.elapsed = function () {
    var c = Core.state.clock;
    if (c.running && c.base !== null) { return performance.now() - c.base; }
    if (c.elapsedAtFinish !== null) { return c.elapsedAtFinish; }
    return 0;
  };

  Core.startClock = function () {
    var c = Core.state.clock;
    c.base = performance.now();
    c.elapsedAtFinish = null;
    c.running = true;
  };

  Core.stopClock = function () {
    var c = Core.state.clock;
    if (c.running) { c.elapsedAtFinish = performance.now() - c.base; }
    c.running = false;
  };

  Core.resetClock = function () {
    var c = Core.state.clock;
    c.base = null;
    c.running = false;
    c.elapsedAtFinish = null;
  };

  /* -------------------------------------------------- pruebas */
  Core.createTest = function (name, teams, starterIds) {
    var starters = starterIds || [];
    var test = {
      id: Core.uid(),
      nombre: name || ('Prueba ' + (Core.state.history.length + 1)),
      startedAt: new Date().toISOString(),
      finishedAt: null,
      status: 'running',
      elapsedAtFinish: null,
      order: [],
      entries: {},
      tieBreaks: {}
    };
    teams.forEach(function (t) {
      test.order.push(t.id);
      test.entries[t.id] = {
        id: t.id,
        nombre: t.nombre,
        int1: t.int1 || '',
        int2: t.int2 || '',
        start: (starters.indexOf(t.id) !== -1) ? 0 : null,
        finish: null,
        duration: null,
        position: null,
        dnf: false
      };
    });
    return test;
  };

  Core.testTeams = function (test) {
    if (!test) { return []; }
    return test.order.map(function (id) { return test.entries[id]; }).filter(Boolean);
  };

  Core.markStart = function (ids) {
    return Core.markStartAt(ids, Core.elapsed());
  };

  Core.markStartAt = function (ids, timeMs) {
    var test = Core.state.current;
    if (!test) { return 0; }
    var n = 0;
    ids.forEach(function (id) {
      var e = test.entries[id];
      if (e && e.start === null && e.finish === null) { e.start = timeMs; n++; }
    });
    Core.recomputePositions(test);
    return n;
  };

  Core.markFinish = function (ids) {
    return Core.markFinishAt(ids, Core.elapsed());
  };

  Core.markFinishAt = function (ids, timeMs) {
    var test = Core.state.current;
    if (!test) { return 0; }
    var n = 0;
    ids.forEach(function (id) {
      var e = test.entries[id];
      if (e && e.finish === null && e.start !== null) {
        e.finish = timeMs;
        e.duration = Math.max(0, timeMs - e.start);
        n++;
      }
    });
    Core.recomputePositions(test);
    return n;
  };

  // Posicion final por duracion (menor = mejor). Empates comparten posicion
  // salvo que exista un desempate manual (test.tieBreaks).
  Core.recomputePositions = function (test) {
    if (!test) { return; }
    var tb = test.tieBreaks || {};
    var finished = Core.testTeams(test).filter(function (e) { return e.finish !== null && e.duration !== null; });
    finished.sort(function (a, b) {
      if (a.duration !== b.duration) { return a.duration - b.duration; }
      var ra = tb[a.id], rb = tb[b.id];
      if (ra != null && rb != null && ra !== rb) { return ra - rb; }
      if (ra != null && rb == null) { return -1; }
      if (rb != null && ra == null) { return 1; }
      var fa = a.finish === null ? Infinity : a.finish;
      var fb = b.finish === null ? Infinity : b.finish;
      if (fa !== fb) { return fa - fb; }
      return a.nombre.localeCompare(b.nombre, 'es');
    });
    var groupStart = 0, i = 0;
    while (i < finished.length) {
      var j = i;
      while (j < finished.length && finished[j].duration === finished[i].duration) { j++; }
      var group = finished.slice(i, j);
      var broken = group.every(function (e) { return tb[e.id] != null; });
      group.forEach(function (e, k) { e.position = broken ? (groupStart + k + 1) : (groupStart + 1); });
      groupStart += group.length;
      i = j;
    }
    Core.testTeams(test).forEach(function (e) {
      if (e.finish === null) { e.position = null; }
    });
  };

  /* -------------------------------------------------- empates y desempate */
  // Grupos de equipos que terminan con exactamente la misma duracion
  Core.findTies = function (test) {
    var finished = Core.testTeams(test).filter(function (e) { return e.finish !== null && e.duration !== null; });
    var groups = [], byDur = {};
    finished.forEach(function (e) {
      var k = String(e.duration);
      if (!byDur[k]) { byDur[k] = []; groups.push(byDur[k]); }
      byDur[k].push(e);
    });
    return groups.filter(function (g) { return g.length > 1; });
  };

  // true si ese grupo ya tiene desempate manual
  Core.tieResolved = function (test, group) {
    var tb = (test && test.tieBreaks) || {};
    return group.every(function (e) { return tb[e.id] != null; });
  };

  // Aplica el orden manual de un grupo empatado (ids en orden de mejor a peor)
  Core.applyTieOrder = function (test, orderedIds) {
    if (!test || !orderedIds || !orderedIds.length) { return; }
    if (!test.tieBreaks) { test.tieBreaks = {}; }
    orderedIds.forEach(function (id, i) { test.tieBreaks[id] = i; });
    Core.recomputePositions(test);
  };

  // Vuelve a dejar el empate sin resolver
  Core.clearTieOrder = function (test, group) {
    if (!test || !test.tieBreaks) { return; }
    group.forEach(function (e) { delete test.tieBreaks[e.id]; });
    Core.recomputePositions(test);
  };

  Core.finalizeTest = function () {
    var test = Core.state.current;
    if (!test) { return null; }
    Core.stopClock();
    Core.recomputePositions(test);
    Core.testTeams(test).forEach(function (e) {
      if (e.finish === null) { e.dnf = true; }
    });
    test.status = 'finished';
    test.finishedAt = new Date().toISOString();
    test.elapsedAtFinish = Core.elapsed();
    Core.state.history.unshift(test);
    Core.state.current = null;
    return test;
  };

  // Estado visual de una entrada
  Core.entryStatus = function (e) {
    if (e.finish !== null) { return 'finished'; }
    if (e.start !== null) { return 'running'; }
    return 'pending';
  };

  App.Core = Core;
})(window.App);
