/* =========================================================
   excel.js - importacion de equipos desde .xlsx / .xls / .csv
   ========================================================= */
window.App = window.App || {};

(function (App) {
  'use strict';

  var Excel = {};

  /* -------------------------------------------------- helpers CSV */
  function stripBom(text) { return text.replace(/^\uFEFF/, ''); }

  function detectDelimiter(text) {
    var line = (stripBom(text).split(/\r?\n/).filter(function (l) { return l.trim() !== ''; })[0] || '');
    var counts = { ',': 0, ';': 0, '\t': 0, '|': 0 };
    var inQ = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i];
      if (ch === '"') { inQ = !inQ; continue; }
      if (inQ) { continue; }
      if (counts[ch] !== undefined) { counts[ch]++; }
    }
    var best = ',', bestCount = -1;
    Object.keys(counts).forEach(function (k) {
      if (counts[k] > bestCount) { bestCount = counts[k]; best = k; }
    });
    return bestCount <= 0 ? ',' : best;
  }

  function parseDelimited(text, delim) {
    var rows = [], row = [], field = '', i = 0, inQ = false;
    text = stripBom(text);
    while (i < text.length) {
      var ch = text[i];
      if (inQ) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        field += ch; i++; continue;
      }
      if (ch === '"') { inQ = true; i++; continue; }
      if (ch === delim) { row.push(field); field = ''; i++; continue; }
      if (ch === '\r') { i++; continue; }
      if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += ch; i++;
    }
    if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
    return rows;
  }

  /* -------------------------------------------------- cabeceras */
  function matchHeader(norm) {
    if (!norm) { return null; }
    if (norm.indexOf('integrante') !== -1 || norm.indexOf('miembro') !== -1) {
      if (/1$|^1|uno|primer|primero/.test(norm)) { return 'int1'; }
      if (/2$|^2|dos|segund/.test(norm)) { return 'int2'; }
      return null;
    }
    if (norm.indexOf('equipo') !== -1 || norm.indexOf('team') !== -1 || norm.indexOf('plantel') !== -1) { return 'equipo'; }
    if (norm === 'nombre' || norm === 'name') { return 'equipo'; }
    return null;
  }

  function findHeaderRow(rows) {
    var limit = Math.min(rows.length, 6);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [];
      var hits = 0;
      for (var c = 0; c < row.length; c++) {
        if (matchHeader(App.Core.normalize(row[c]))) { hits++; }
      }
      if (hits >= 1) { return r; }
    }
    return -1;
  }

  /* -------------------------------------------------- parseo de filas */
  Excel.parseRows = function (rows) {
    var result = { teams: [], warnings: [], mapping: {}, skipped: 0, headerRow: -1 };
    rows = (rows || []).filter(function (r) {
      return r && r.some(function (c) { return String(c == null ? '' : c).trim() !== ''; });
    });

    if (!rows.length) {
      result.warnings.push('El archivo esta vacio.');
      return result;
    }

    var headerRow = findHeaderRow(rows);
    var map = { equipo: -1, int1: -1, int2: -1 };
    var dataStart = 0;

    if (headerRow >= 0) {
      dataStart = headerRow + 1;
      var header = rows[headerRow] || [];
      for (var c = 0; c < header.length; c++) {
        var hit = matchHeader(App.Core.normalize(header[c]));
        if (hit && map[hit] === -1) { map[hit] = c; }
      }
      result.headerRow = headerRow;
      if (map.equipo === -1 && map.int1 === -1 && map.int2 === -1) {
        headerRow = -1; dataStart = 0;
      } else if (dataStart >= rows.length) {
        // la fila detectada era la ultima de datos, no una cabecera
        headerRow = -1; dataStart = 0;
        result.headerRow = -1;
        result.warnings.push('No se encontro una fila de encabezados: se usaron las columnas A, B y C.');
      }
    }

    if (headerRow < 0) {
      map = { equipo: 0, int1: 1, int2: 2 };
      if (!result.warnings.length) {
        result.warnings.push('No se reconocieron las cabeceras: se usaron las columnas A (Nombre Equipo), B (Integrante 1) y C (Integrante 2).');
      }
    } else {
      if (map.equipo === -1) {
        map.equipo = 0;
        result.warnings.push('No se encontro la columna "Nombre Equipo": se uso la columna A.');
      }
      if (map.int1 === -1) { map.int1 = (map.equipo === 1 ? 2 : 1); }
      if (map.int2 === -1) { map.int2 = (map.equipo === 2 ? 1 : map.equipo + 2); }
    }

    result.mapping = map;

    var seen = {};
    for (var r = dataStart; r < rows.length; r++) {
      var row = rows[r] || [];
      var nombre = String(row[map.equipo] == null ? '' : row[map.equipo]).replace(/\s+/g, ' ').trim();
      if (!nombre) { result.skipped++; continue; }
      var key = nombre.toLowerCase();
      if (seen[key]) {
        result.skipped++;
        result.warnings.push('Equipo duplicado ignorado: "' + nombre + '".');
        continue;
      }
      seen[key] = true;
      result.teams.push({
        nombre: nombre,
        int1: String(row[map.int1] == null ? '' : row[map.int1]).replace(/\s+/g, ' ').trim(),
        int2: String(row[map.int2] == null ? '' : row[map.int2]).replace(/\s+/g, ' ').trim()
      });
    }

    if (!result.teams.length) {
      result.warnings.push('No se encontro ningun equipo en el archivo.');
    }
    return result;
  };

  /* -------------------------------------------------- lectura de archivo */
  Excel.readFile = function (file) {
    return new Promise(function (resolve, reject) {
      if (!file) { reject(new Error('Archivo vacio.')); return; }
      var name = (file.name || '').toLowerCase();
      var isCsv = /\.(csv|txt|tsv)$/.test(name);

      if (isCsv) {
        var reader = new FileReader();
        reader.onerror = function () { reject(new Error('No se pudo leer el archivo.')); };
        reader.onload = function (e) {
          try {
            var text = String(e.target.result || '');
            var delim = detectDelimiter(text);
            resolve(Excel.parseRows(parseDelimited(text, delim)));
          } catch (err) {
            reject(err);
          }
        };
        reader.readAsText(file, 'utf-8');
        return;
      }

      var reader2 = new FileReader();
      reader2.onerror = function () { reject(new Error('No se pudo leer el archivo.')); };
      reader2.onload = function (e) {
        try {
          if (typeof XLSX === 'undefined') { throw new Error('Libreria de Excel no disponible.'); }
          var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
          var sheetName = wb.SheetNames[0];
          if (!sheetName) { throw new Error('El archivo no tiene hojas.'); }
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false, defval: '' });
          var parsed = Excel.parseRows(rows);
          parsed.sheet = sheetName;
          resolve(parsed);
        } catch (err) {
          reject(err);
        }
      };
      reader2.readAsArrayBuffer(file);
    });
  };

  App.Excel = Excel;
})(window.App);
