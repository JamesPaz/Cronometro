/* =========================================================
   storage.js - persistencia local + exportacion a Excel/CSV/JSON
   ========================================================= */
window.App = window.App || {};

(function (App) {
  'use strict';

  var Core = App.Core;
  var Storage = {};

  var KEY = 'cronometro.v1';
  var DELIM = ';';               // separador compatible con Excel en español
  var HEADERS = ['Nombre Equipo', 'Tiempo Inicio', 'Tiempo Fin', 'Duración', 'Posición', 'Integrante 1', 'Integrante 2'];

  /* -------------------------------------------------- localStorage */
  Storage.KEY = KEY;

  Storage.saveNow = function (state) {
    try {
      localStorage.setItem(KEY, JSON.stringify({
        v: 1,
        roster: state.roster,
        current: state.current,
        history: state.history,
        clockElapsed: Core.elapsed(),
        ui: state.ui || { collapsed: [] },
        savedAt: new Date().toISOString()
      }));
      return true;
    } catch (e) {
      return false;
    }
  };

  var timer = null;
  Storage.save = function (state) {
    if (timer) { clearTimeout(timer); }
    timer = setTimeout(function () { Storage.saveNow(state); }, 350);
  };

  Storage.load = function () {
    try {
      var raw = localStorage.getItem(KEY);
      if (!raw) { return null; }
      var data = JSON.parse(raw);
      if (!data || typeof data !== 'object') { return null; }
      return {
        roster: data.roster || { teams: [] },
        current: data.current || null,
        history: data.history || [],
        clockElapsed: typeof data.clockElapsed === 'number' ? data.clockElapsed : null,
        ui: data.ui || { collapsed: [] }
      };
    } catch (e) {
      return null;
    }
  };

  Storage.clear = function () {
    try { localStorage.removeItem(KEY); } catch (e) { /* noop */ }
  };

  /* -------------------------------------------------- descargas */
  Storage.downloadBlob = function (filename, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  };

  function csvEscape(value) {
    var s = (value == null ? '' : String(value));
    if (s.indexOf(DELIM) !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  function toCSV(header, rows) {
    var lines = [header.map(csvEscape).join(DELIM)];
    rows.forEach(function (r) { lines.push(r.map(csvEscape).join(DELIM)); });
    return '\uFEFF' + lines.join('\r\n') + '\r\n';
  }

  /* -------------------------------------------------- filas de resultado */
  function entryRow(e) {
    var pos;
    if (e.position != null) { pos = String(e.position); }
    else if (e.start === null) { pos = 'No inicio'; }
    else { pos = 'No termino'; }
    return [
      e.nombre,
      e.start === null ? '' : Core.fmtFull(e.start),
      e.finish === null ? '' : Core.fmtFull(e.finish),
      e.duration === null ? '' : Core.fmtFull(e.duration),
      pos,
      e.int1 || '',
      e.int2 || ''
    ];
  }

  function sortedEntries(test) {
    return Core.testTeams(test).slice().sort(function (a, b) {
      var pa = a.position != null ? a.position : 9999;
      var pb = b.position != null ? b.position : 9999;
      if (pa !== pb) { return pa - pb; }
      return a.nombre.localeCompare(b.nombre, 'es');
    });
  }

  /* -------------------------------------------------- exportaciones */
  function safeName(str) {
    return String(str == null || str === '' ? 'prueba' : str)
      .replace(/[\\/:*?"<>|\r\n\t]+/g, ' ')
      .trim()
      .replace(/\s+/g, '_')
      .slice(0, 48) || 'prueba';
  }

  Storage.exportResultsCSV = function (test) {
    var rows = sortedEntries(test).map(entryRow);
    var csv = toCSV(HEADERS, rows);
    Storage.downloadBlob('resultados_' + safeName(test.nombre) + '_' + Core.fmtFileStamp(test.startedAt) + '.csv',
      new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
  };

  Storage.exportRosterCSV = function (teams) {
    var rows = teams.map(function (t) { return [t.nombre, t.int1 || '', t.int2 || '']; });
    var csv = toCSV(['Nombre Equipo', 'Integrante 1', 'Integrante 2'], rows);
    Storage.downloadBlob('equipos_' + Core.fmtFileStamp(new Date().toISOString()) + '.csv',
      new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
  };

  Storage.exportHistoryCSV = function (history) {
    var rows = [];
    history.forEach(function (test) {
      sortedEntries(test).forEach(function (e) {
        var r = entryRow(e);
        rows.push([test.nombre, Core.fmtDateTime(test.startedAt)].concat(r));
      });
    });
    var csv = toCSV(['Prueba', 'Fecha'].concat(HEADERS), rows);
    Storage.downloadBlob('historial_' + Core.fmtFileStamp(new Date().toISOString()) + '.csv',
      new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
  };

  Storage.exportXLSX = function (test, teams) {
    if (typeof XLSX === 'undefined') { throw new Error('Libreria de Excel no disponible.'); }
    var wb = XLSX.utils.book_new();

    var resRows = [HEADERS].concat(sortedEntries(test).map(entryRow));
    var wsRes = XLSX.utils.aoa_to_sheet(resRows);
    wsRes['!cols'] = [{ wch: 26 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 20 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, wsRes, 'Resultados');

    if (teams && teams.length) {
      var teamRows = [['Nombre Equipo', 'Integrante 1', 'Integrante 2']].concat(
        teams.map(function (t) { return [t.nombre, t.int1 || '', t.int2 || '']; }));
      var wsTeam = XLSX.utils.aoa_to_sheet(teamRows);
      wsTeam['!cols'] = [{ wch: 26 }, { wch: 20 }, { wch: 20 }];
      XLSX.utils.book_append_sheet(wb, wsTeam, 'Equipos');
    }

    var metaRows = [
      ['Campo', 'Valor'],
      ['Prueba', test.nombre || ''],
      ['Fecha', Core.fmtDateTime(test.startedAt)],
      ['Equipos', Core.testTeams(test).length],
      ['Finalizada', test.finishedAt ? Core.fmtDateTime(test.finishedAt) : 'En curso']
    ];
    var wsMeta = XLSX.utils.aoa_to_sheet(metaRows);
    wsMeta['!cols'] = [{ wch: 16 }, { wch: 38 }];
    XLSX.utils.book_append_sheet(wb, wsMeta, 'Prueba');

    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    var blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    Storage.downloadBlob('resultados_' + safeName(test.nombre) + '_' + Core.fmtFileStamp(test.startedAt) + '.xlsx', blob);
  };

  Storage.exportJSON = function (state) {
    var data = {
      app: 'cronometro-equipos',
      v: 1,
      exportedAt: new Date().toISOString(),
      roster: state.roster,
      current: state.current,
      history: state.history,
      clockElapsed: Core.elapsed(),
      ui: state.ui || { collapsed: [] }
    };
    var blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json;charset=utf-8;' });
    Storage.downloadBlob('cronometro_respaldo_' + Core.fmtFileStamp(new Date().toISOString()) + '.json', blob);
  };

  Storage.importJSON = function (file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('No se pudo leer el archivo.')); };
      reader.onload = function (e) {
        try {
          var data = JSON.parse(String(e.target.result || '{}'));
          if (!data || typeof data !== 'object') { throw new Error('Archivo no valido.'); }
          resolve({
            roster: data.roster || { teams: [] },
            current: data.current || null,
            history: data.history || [],
            clockElapsed: typeof data.clockElapsed === 'number' ? data.clockElapsed : null,
            ui: data.ui || { collapsed: [] }
          });
        } catch (err) {
          reject(new Error('El archivo JSON no es valido.'));
        }
      };
      reader.readAsText(file, 'utf-8');
    });
  };

  App.Storage = Storage;
})(window.App);
