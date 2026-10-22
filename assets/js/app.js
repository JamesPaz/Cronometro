/* =========================================================
   app.js - interfaz, eventos y flujo de la prueba
   ========================================================= */
(function ($) {
  'use strict';

  var App = window.App;
  var Core = App.Core;
  var Excel = App.Excel;
  var Storage = App.Storage;

  var $ = jQuery;
  var els = {};
  var viewTest = null;          // prueba mostrada en el podio (no persistido)
  var pickerState = null;       // estado del selector de equipos
  var pendingExcelAction = null;
  var lastSaveAt = null;

  function save() {
    var ok = Storage.saveNow(Core.state);
    if (ok) {
      var d = new Date();
      lastSaveAt = Core.pad(d.getHours(), 2) + ':' + Core.pad(d.getMinutes(), 2) + ':' + Core.pad(d.getSeconds(), 2);
    }
    if (els.saveState) {
      els.saveState.text(ok
        ? 'Guardado automatico activo - ultimo guardado ' + lastSaveAt
        : 'No se pudo guardar en este navegador (revisa el modo privado).');
    }
  }

  /* =====================================================
     toasts
     ===================================================== */
  function toast(msg, type, ms) {
    var $t = $('<div class="toast"></div>')
      .addClass('toast--' + (type || 'ok'))
      .text(msg)
      .appendTo(els.toasts);
    setTimeout(function () {
      $t.addClass('is-out');
      setTimeout(function () { $t.remove(); }, 350);
    }, ms || 3600);
  }

  /* =====================================================
     modal generico
     ===================================================== */
  var modalGen = 0;

  function openModal(cfg) {
    cfg = cfg || {};
    var myGen = ++modalGen;
    els.modalTitle.text(cfg.title || '');
    els.modalSub.html(cfg.sub || '');
    els.modalTools.toggle(cfg.showTools !== false);
    els.modalSearch.val('');
    els.modalBody.html(cfg.html || '');

    var buttons = cfg.buttons || [{ label: 'Confirmar', cls: 'btn btn--primary', value: 'ok' }];
    els.modalFoot.empty();
    buttons.forEach(function (b) {
      $('<button></button>').addClass(b.cls || 'btn btn--ghost').text(b.label).attr('data-value', b.value).appendTo(els.modalFoot);
    });

    els.modal.off('click.m').on('click.m', '[data-value]', function () {
      var value = $(this).attr('data-value');
      if (value === 'ok') {
        var res = cfg.onConfirm ? cfg.onConfirm(els.modalBody) : true;
        if (res === false) { return; }
        if (modalGen === myGen) { closeModal(); }
      } else if (cfg.onButton) {
        var r2 = cfg.onButton(value, els.modalBody);
        if (r2 === false) { return; }
        if (modalGen === myGen) { closeModal(); }
      } else if (modalGen === myGen) {
        closeModal();
      }
    });

    els.modal.removeAttr('hidden');
    if (cfg.onOpen) { cfg.onOpen(els.modalBody); }
  }

  function closeModal() {
    els.modal.attr('hidden', true);
    modalGen++;
    pickerState = null;
  }

  /* =====================================================
     tarjetas colapsables del panel
     ===================================================== */
  function collapsedList() {
    if (!Core.state.ui) { Core.state.ui = { collapsed: [] }; }
    if (!Array.isArray(Core.state.ui.collapsed)) { Core.state.ui.collapsed = []; }
    return Core.state.ui.collapsed;
  }

  function isCollapsed(id) {
    return collapsedList().indexOf(id) !== -1;
  }

  function setCollapsed(id, value) {
    var list = collapsedList();
    var i = list.indexOf(id);
    if (value && i === -1) { list.push(id); }
    else if (!value && i !== -1) { list.splice(i, 1); }
    save();
  }

  function applyCollapsed() {
    els.panel.find('.card[data-card]').each(function () {
      var $card = $(this);
      var id = $card.attr('data-card');
      var off = isCollapsed(id);
      $card.toggleClass('is-collapsed', off);
      $card.find('.card__head').first().attr('aria-expanded', off ? 'false' : 'true');
    });
  }

  function toggleCard(id) {
    setCollapsed(id, !isCollapsed(id));
    applyCollapsed();
  }

  function initCollapsible() {
    els.panel.find('.card[data-card]').each(function () {
      var $card = $(this);
      var id = $card.attr('data-card');
      var $head = $card.find('.card__head').first();
      if (!$head.length) { return; }
      $head.attr('role', 'button').attr('tabindex', '0');
      if (!$head.find('.card__chev').length) {
        $head.append('<span class="card__chev" aria-hidden="true"></span>');
      }
      $head.off('click.card').on('click.card', function (e) {
        if ($(e.target).closest('button, input, select, textarea, a').length) { return; }
        e.preventDefault();
        toggleCard(id);
      });
      $head.off('keydown.card').on('keydown.card', function (e) {
        if (e.key === 'Enter' || e.key === ' ' || e.key === 'Spacebar') {
          e.preventDefault();
          toggleCard(id);
        }
      });
    });
  }

  /* =====================================================
     selector de equipos (modal con checkboxes)
     ===================================================== */
  function memberLabel(t) {
    var parts = [];
    if (t.int1) { parts.push(t.int1); }
    if (t.int2) { parts.push(t.int2); }
    return parts.join(' - ');
  }

  function openPicker(cfg) {
    var state = { items: [], checked: {} };
    function readItems() {
      return cfg.getItems ? cfg.getItems() : (cfg.items || []);
    }
    function defaultFor(item) {
      return cfg.defaultChecked !== undefined ? cfg.defaultChecked : !!item.checked;
    }
    state.items = readItems();
    state.items.forEach(function (it) { state.checked[it.id] = defaultFor(it); });
    pickerState = state;

    function renderList() {
      var items = state.items;
      var html = cfg.extraTopHtml ? cfg.extraTopHtml() : '';
      if (!items.length) {
        html += '<div class="sel__empty">' + (cfg.emptyText || 'No hay equipos disponibles.') + '</div>';
      } else {
        html += '<div class="sel">';
        items.forEach(function (it) {
          var on = !!state.checked[it.id];
          html += '<label class="sel__item' + (on ? ' is-checked' : '') + '" data-search="' +
            Core.escape((it.nombre + ' ' + (it.meta || '')).toLowerCase()) + '">' +
            '<input type="checkbox" data-id="' + Core.escape(it.id) + '"' + (on ? ' checked' : '') + '>' +
            '<div class="sel__txt">' +
            '<div class="sel__name">' + Core.escape(it.nombre) + '</div>' +
            (it.meta ? '<div class="sel__meta">' + Core.escape(it.meta) + '</div>' : '') +
            '</div>' +
            (it.tag ? '<div class="sel__tag' + (it.tagGrey ? ' sel__tag--grey' : '') + '">' + Core.escape(it.tag) + '</div>' : '') +
            '</label>';
        });
        html += '</div>';
      }
      return html;
    }

    function refresh() {
      state.items = readItems();
      state.items.forEach(function (it) {
        if (state.checked[it.id] === undefined) { state.checked[it.id] = defaultFor(it); }
      });
      els.modalBody.html(renderList());
      applyFilter();
      if (cfg.onRefresh) { cfg.onRefresh(); }
    }
    state.refresh = refresh;

    function applyFilter() {
      var q = (els.modalSearch.val() || '').toLowerCase().trim();
      els.modalBody.find('.sel__item').each(function () {
        var $i = $(this);
        var hit = !q || $i.attr('data-search').indexOf(q) !== -1;
        $i.toggle(hit);
      });
    }

    function setAll(value) {
      els.modalBody.find('input[type=checkbox]').each(function () {
        var id = $(this).attr('data-id');
        this.checked = value;
        state.checked[id] = value;
        $(this).closest('.sel__item').toggleClass('is-checked', value);
      });
    }

    openModal({
      title: cfg.title,
      sub: cfg.sub,
      showTools: true,
      buttons: cfg.buttons || [{ label: cfg.confirmLabel || 'Confirmar', cls: 'btn btn--primary', value: 'ok' }],
      onConfirm: function () {
        var ids = [];
        els.modalBody.find('input[type=checkbox]:checked').each(function () {
          ids.push($(this).attr('data-id'));
        });
        if (cfg.onConfirm) { return cfg.onConfirm(ids); }
        return true;
      },
      onOpen: function () {
        els.modalBody.off('change.picker').on('change.picker', 'input[type=checkbox]', function () {
          var id = $(this).attr('data-id');
          state.checked[id] = this.checked;
          $(this).closest('.sel__item').toggleClass('is-checked', this.checked);
        });
        els.modalBody.off('input.picker').on('input.picker', '[data-picker-input]', function () {
          if (cfg.onInput) { cfg.onInput(this.value, this); }
        });
        els.modalBody.off('keydown.picker').on('keydown.picker', 'input[type=text]', function (e) {
          if (e.key === 'Enter') {
            e.preventDefault();
            var $ok = els.modalFoot.find('[data-value="ok"]');
            if ($ok.length) { $ok.trigger('click'); }
          }
        });
        els.modalBody.off('click.picker').on('click.picker', '[data-modal-load]', function () {
          pendingExcelAction = function (teams) {
            Core.state.roster.teams = teams;
            save();
            teams.forEach(function (t) { state.checked[t.id] = true; });
            refresh();
            toast('Equipos cargados: ' + teams.length, 'ok');
          };
          els.fileInput.val('').trigger('click');
        });
        refresh();
      }
    });

    els.modalSearch.off('input.picker').on('input.picker', applyFilter);
    els.modalAll.off('click.picker').on('click.picker', function () { setAll(true); });
    els.modalNone.off('click.picker').on('click.picker', function () { setAll(false); });
  }

  /* =====================================================
     carga de equipos
     ===================================================== */
  function toTeams(list) {
    return list.map(function (t) {
      return { id: Core.uid(), nombre: t.nombre, int1: t.int1 || '', int2: t.int2 || '' };
    });
  }

  function handleExcelFile(file) {
    if (!file) { return; }
    Excel.readFile(file).then(function (res) {
      var teams = toTeams(res.teams);
      (res.warnings || []).slice(0, 3).forEach(function (w) { toast(w, 'warn', 5200); });

      if (pendingExcelAction) {
        var action = pendingExcelAction;
        pendingExcelAction = null;
        action(teams);
        return;
      }
      if (!teams.length) {
        toast('No se encontraron equipos en el archivo.', 'err');
        return;
      }
      var current = Core.state.roster.teams;
      if (!current.length) {
        Core.state.roster.teams = teams;
        save();
        render();
        if (pickerState && pickerState.refresh) { pickerState.refresh(); }
        toast('Se cargaron ' + teams.length + ' equipos.', 'ok');
        return;
      }
      openModal({
        title: 'Ya hay ' + current.length + ' equipos cargados',
        sub: 'Elige que hacer con el archivo <b>' + Core.escape(file.name) + '</b>.',
        showTools: false,
        html: '<p class="modal__msg">El archivo contiene <b>' + teams.length + '</b> equipos.</p>',
        buttons: [
          { label: 'Reemplazar', cls: 'btn btn--primary', value: 'replace' },
          { label: 'Agregar', cls: 'btn btn--ghost', value: 'merge' },
          { label: 'Cancelar', cls: 'btn btn--ghost', value: 'cancel' }
        ],
        onButton: function (value) {
          if (value === 'replace') {
            Core.state.roster.teams = teams;
          } else if (value === 'merge') {
            var existing = {};
            Core.state.roster.teams.forEach(function (t) { existing[t.nombre.toLowerCase()] = true; });
            teams.forEach(function (t) {
              if (!existing[t.nombre.toLowerCase()]) { Core.state.roster.teams.push(t); }
            });
          } else { return; }
          save();
          render();
          toast('Lista de equipos actualizada (' + Core.state.roster.teams.length + ').', 'ok');
        }
      });
    }).catch(function (err) {
      toast('Error al leer el archivo: ' + (err && err.message ? err.message : 'desconocido'), 'err', 5200);
    });
  }

  /* =====================================================
     desempate manual
     ===================================================== */
  function orderedTieGroup(test, group) {
    var tb = test.tieBreaks || {};
    return group.slice().sort(function (a, b) {
      var ra = tb[a.id], rb = tb[b.id];
      if (ra != null && rb != null && ra !== rb) { return ra - rb; }
      if (ra != null && rb == null) { return -1; }
      if (rb != null && ra == null) { return 1; }
      if (a.finish !== b.finish) { return a.finish - b.finish; }
      return a.nombre.localeCompare(b.nombre, 'es');
    });
  }

  function tieGroupHtml(test, group, gi) {
    var resolved = Core.tieResolved(test, group);
    var html = '<div class="tie' + (resolved ? '' : '') + '" data-group="' + gi + '" data-dur="' + group[0].duration + '"' +
      (resolved ? ' data-resolved="1"' : '') + '>';
    html += '<div class="tie__head">Empate en <b>' + Core.fmtFull(group[0].duration) + '</b> &middot; ' +
      group.length + ' equipos' + (resolved ? ' &middot; <b style="color:var(--green)">desempate aplicado</b>' : '') + '</div>';
    html += '<ul class="tie__list">';
    orderedTieGroup(test, group).forEach(function (e, i) {
      html += '<li class="tie__item" data-id="' + Core.escape(e.id) + '">' +
        '<span class="tie__rank">' + (i + 1) + '</span>' +
        '<span class="tie__name">' + Core.escape(e.nombre) + '</span>' +
        '<span class="tie__meta">' + Core.escape(memberLabel(e)) + '</span>' +
        '<button class="iconbtn" data-tie="up" title="Subir">&uarr;</button>' +
        '<button class="iconbtn" data-tie="down" title="Bajar">&darr;</button>' +
        '</li>';
    });
    html += '</ul>';
    html += '<button class="btn btn--ghost btn--xs tie__reset" data-tie-reset="1">Restaurar empate (posicion compartida)</button>';
    html += '</div>';
    return html;
  }

  function renumberTie($list) {
    $list.find('.tie__item').each(function (i) {
      $(this).find('.tie__rank').text(String(i + 1));
    });
  }

  function openTieBreak(test) {
    test = test || Core.state.current;
    if (!test) { toast('No hay prueba para resolver.', 'warn'); return; }
    var groups = Core.findTies(test);
    if (!groups.length) {
      toast('No hay empates por resolver.', 'ok');
      return;
    }
    var pending = groups.filter(function (g) { return !Core.tieResolved(test, g); }).length;

    var html = '<p class="modal__msg">Define manualmente <b>quien queda por encima</b> en cada empate. ' +
      'Los <b>tiempos no cambian</b>: solo se define el orden final y las posiciones. ' +
      (pending ? 'Hay <b>' + pending + '</b> empate(s) sin resolver.' : 'Todos los empates ya estan resueltos.') +
      '</p>';
    groups.forEach(function (g, gi) { html += tieGroupHtml(test, g, gi); });

    openModal({
      title: 'Resolver empates',
      sub: 'Prueba: <b>' + Core.escape(test.nombre) + '</b>',
      showTools: false,
      html: html,
      buttons: [
        { label: 'Cancelar', cls: 'btn btn--ghost', value: 'cancel' },
        { label: 'Aplicar desempate', cls: 'btn btn--primary', value: 'ok' }
      ],
      onConfirm: function () {
        var fresh = Core.findTies(test);
        els.modalBody.find('.tie').each(function () {
          var $tie = $(this);
          var dur = Number($tie.attr('data-dur'));
          var group = fresh.filter(function (g) { return g[0].duration === dur; })[0];
          if (!group) { return; }
          if ($tie.attr('data-reset') === '1') {
            Core.clearTieOrder(test, group);
            return;
          }
          var ids = $tie.find('.tie__item').map(function () { return $(this).attr('data-id'); }).get();
          Core.applyTieOrder(test, ids);
        });
        save();
        render();
        toast('Desempate aplicado. Posiciones actualizadas.', 'ok');
        return true;
      },
      onOpen: function (body) {
        body.off('.tie')
          .on('click.tie', '[data-tie]', function () {
            var dir = $(this).attr('data-tie');
            var $item = $(this).closest('.tie__item');
            var $list = $item.parent();
            var $sib = dir === 'up' ? $item.prev() : $item.next();
            if (!$sib.length) { return; }
            if (dir === 'up') { $item.insertBefore($sib); } else { $item.insertAfter($sib); }
            renumberTie($list);
          })
          .on('click.tie', '[data-tie-reset]', function () {
            var $tie = $(this).closest('.tie');
            $tie.attr('data-reset', '1').addClass('is-reset');
            $(this).text('Se restaurara el empate al aplicar');
          });
      }
    });
  }

  /* =====================================================
     flujo de la prueba
     ===================================================== */

  // Configuracion: nombre de la prueba + elegir equipos que arrancan en 0
  var setupName = '';

  function nameFieldHtml() {
    return '<div class="field field--block">' +
      '<label class="field__label" for="testNameInput">Nombre de la prueba</label>' +
      '<input type="text" id="testNameInput" class="input" data-picker-input="1" maxlength="60" ' +
      'placeholder="Ej: Serie A - Final" value="' + Core.escape(setupName) + '">' +
      '</div>';
  }

  function openSetup() {
    setupName = nextTestName();

    openPicker({
      title: 'Configurar prueba',
      sub: 'Ponle nombre a la prueba y marca los equipos que <b>arrancan en 00:00:00</b>. Los que dejes sin marcar podras iniciarlos despues con "Marcar inicio".',
      confirmLabel: 'Empezar a cronometrar',
      defaultChecked: true,
      items: [],
      extraTopHtml: function () {
        return nameFieldHtml() +
          '<div class="row" style="margin-bottom:12px">' +
          '<button class="btn btn--primary btn--sm" data-modal-load="1">Cargar Excel (.xlsx / .csv)</button>' +
          '<span class="note note--dim">Equipos cargados: ' + Core.state.roster.teams.length + '</span>' +
          '</div>';
      },
      getItems: function () {
        return Core.state.roster.teams.map(function (t) {
          return { id: t.id, nombre: t.nombre, meta: memberLabel(t), checked: true };
        });
      },
      onInput: function (v) { setupName = v; },
      onRefresh: function () {
        var n = Core.state.roster.teams.length;
        var $ok = els.modalFoot.find('[data-value="ok"]');
        if (n) {
          els.modalTitle.text('Configurar prueba');
          els.modalSub.html('Ponle nombre a la prueba y marca los equipos que <b>arrancan en 00:00:00</b>. Los que dejes sin marcar podras iniciarlos despues con "Marcar inicio".');
          $ok.text('Empezar a cronometrar');
        } else {
          els.modalTitle.text('Carga los equipos');
          els.modalSub.html('No hay equipos cargados. Importa tu Excel (.xlsx / .csv) con las columnas <b>Nombre Equipo</b>, <b>Integrante 1</b> e <b>Integrante 2</b>.');
          $ok.text('Entendido');
        }
      },
      onConfirm: function (ids) {
        if (!Core.state.roster.teams.length) { return true; }
        if (!ids.length) {
          toast('Selecciona al menos un equipo para arrancar en 0.', 'warn');
          return false;
        }
        var typed = (els.modalBody.find('#testNameInput').val() || '').trim();
        var name = typed || setupName.trim() || nextTestName();
        startTest(ids, name);
        return true;
      }
    });
  }

  function startTest(starterIds, name) {
    var teams = Core.state.roster.teams;
    if (!teams.length) {
      toast('Carga los equipos primero.', 'warn');
      return;
    }
    Core.state.current = Core.createTest((name || '').trim() || nextTestName(), teams, starterIds);
    Core.resetClock();
    Core.startClock();
    viewTest = null;
    save();
    render();
    var n = starterIds.length;
    toast('Prueba "' + Core.state.current.nombre + '" iniciada. ' +
      (n === 1 ? '1 equipo arranca' : n + ' equipos arrancan') + ' en 00:00:00.', 'ok');
  }

  function nextTestName() {
    var n = Core.state.history.length + 1;
    var names = Core.state.history.map(function (t) { return t.nombre; });
    while (names.indexOf('Prueba ' + n) !== -1) { n++; }
    return 'Prueba ' + n;
  }

  function openStartPicker() {
    var test = Core.state.current;
    if (!test) { toast('No hay una prueba en curso.', 'warn'); return; }
    var items = Core.testTeams(test).filter(function (e) { return e.start === null && e.finish === null; })
      .map(function (e) {
        return { id: e.id, nombre: e.nombre, meta: memberLabel(e), tag: 'sin iniciar', tagGrey: true, checked: true };
      });
    if (!items.length) {
      toast('Todos los equipos ya tienen hora de inicio.', 'warn');
      return;
    }
    openPicker({
      title: 'Marcar inicio',
      sub: 'Selecciona los equipos que <b>arrancan ahora</b>. Tiempo actual: <b>' + Core.fmtFull(Core.elapsed()) + '</b>',
      confirmLabel: 'Registrar inicio',
      items: items,
      onConfirm: function (ids) {
        if (!ids.length) { toast('Selecciona al menos un equipo.', 'warn'); return false; }
        var n = Core.markStart(ids);
        save();
        render();
        toast('Inicio registrado para ' + n + (n === 1 ? ' equipo.' : ' equipos.'), 'ok');
      }
    });
  }

  function openFinishPicker() {
    var test = Core.state.current;
    if (!test) { toast('No hay una prueba en curso.', 'warn'); return; }
    var items = Core.testTeams(test).filter(function (e) { return e.start !== null && e.finish === null; })
      .map(function (e) {
        return {
          id: e.id, nombre: e.nombre, meta: memberLabel(e),
          tag: 'en curso ' + Core.fmtFull(Math.max(0, Core.elapsed() - e.start)), checked: true
        };
      });
    if (!items.length) {
      toast('No hay equipos en curso para marcar.', 'warn');
      return;
    }
    openPicker({
      title: 'Marcar fin',
      sub: 'Selecciona los equipos que <b>terminaron la prueba</b>. Tiempo actual: <b>' + Core.fmtFull(Core.elapsed()) + '</b>',
      confirmLabel: 'Registrar fin',
      items: items,
      onConfirm: function (ids) {
        if (!ids.length) { toast('Selecciona al menos un equipo.', 'warn'); return false; }
        var n = Core.markFinish(ids);
        save();
        render();
        toast('Fin registrado para ' + n + (n === 1 ? ' equipo.' : ' equipos.'), 'ok');
      }
    });
  }

  function openFinishTest() {
    var test = Core.state.current;
    if (!test) { toast('No hay una prueba en curso.', 'warn'); return; }
    var pending = Core.testTeams(test).filter(function (e) { return e.finish === null; });
    var ties = Core.findTies(test);
    var html = '<p class="modal__msg">Se calcularan las posiciones finales y se mostrara el <b>podio</b>.' +
      (ties.length
        ? '<p class="note" style="margin-top:10px;color:var(--amber)">Hay <b>' + ties.length + '</b> empate(s) por duracion. ' +
          'Al finalizar podras definir manualmente quien queda por encima.</p>'
        : '') +
      (pending.length
        ? '<ul>' + pending.map(function (e) {
          return '<li>' + Core.escape(e.nombre) + ' - ' + (e.start === null ? 'no arranco' : 'en curso') + '</li>';
        }).join('') + '</ul>'
        : '<p class="note">Todos los equipos ya terminaron.</p>') +
      '</p>';

    openModal({
      title: 'Finalizar prueba',
      sub: 'Prueba en curso: <b>' + Core.escape(test.nombre) + '</b>',
      showTools: false,
      html: html,
      buttons: [
        { label: 'Cancelar', cls: 'btn btn--ghost', value: 'cancel' },
        { label: 'Finalizar y ver podio', cls: 'btn btn--danger', value: 'ok' }
      ],
      onConfirm: function () {
        var finished = Core.finalizeTest();
        viewTest = finished;
        save();
        render();
        var newTies = Core.findTies(finished);
        if (newTies.length) {
          toast('Hay ' + newTies.length + ' empate(s). Define quien queda por encima.', 'warn', 5200);
          setTimeout(function () { openTieBreak(finished); }, 380);
        } else {
          toast('Prueba finalizada. Podio calculado.', 'ok');
        }
      }
    });
  }

  function requestNewTest() {
    var test = Core.state.current;
    if (!test) {
      openSetup();
      return;
    }
    openModal({
      title: 'Empezar otra prueba',
      sub: 'La prueba <b>' + Core.escape(test.nombre) + '</b> esta en curso.',
      showTools: false,
      html: '<p class="modal__msg">Puedes <b>finalizarla</b> (se guardara en el historial con su podio) o <b>descartarla</b>. ' +
        'En ambos casos el cronometro volvera a 00:00:00.</p>',
      buttons: [
        { label: 'Cancelar', cls: 'btn btn--ghost', value: 'cancel' },
        { label: 'Descartar', cls: 'btn btn--ghost', value: 'discard' },
        { label: 'Finalizar y seguir', cls: 'btn btn--primary', value: 'ok' }
      ],
      onConfirm: function () {
        var finished = Core.finalizeTest();
        viewTest = finished;
        save();
        openSetup();
      },
      onButton: function (value) {
        if (value === 'discard') {
          Core.state.current = null;
          Core.resetClock();
          viewTest = null;
          save();
          render();
          openSetup();
        }
      }
    });
  }

  /* =====================================================
     render
     ===================================================== */
  function render() {
    renderStage();
    renderPanel();
    renderLive();
    renderHistory();
    renderButtons();
  }

  function renderStage() {
    var test = Core.state.current;
    var showingPodium = !!viewTest;

    if (showingPodium) {
      els.clockView.hide();
      els.podiumView.removeAttr('hidden');
      renderPodium(viewTest);
      els.testName.text(viewTest.nombre + ' - resultado final');
      els.chipStatus.text('Finalizada').removeClass('is-running').addClass('is-finished');
      els.chipDate.text(Core.fmtDateTime(viewTest.startedAt));
      els.hint.html('Podio final de <b>' + Core.escape(viewTest.nombre) + '</b>. Pulsa <b>Nueva prueba</b> para cargar los equipos y empezar la siguiente.');
    } else {
      els.podiumView.attr('hidden', true);
      els.clockView.show();
      if (test) {
        els.testName.text(test.nombre);
        els.chipStatus.text('En curso').removeClass('is-finished').addClass('is-running');
        els.chipDate.text(Core.fmtDateTime(test.startedAt));
        els.hint.html('El reloj <b>no se detiene</b>: marca el fin de cada equipo para registrar su duracion.');
      } else {
        els.testName.text('Sin prueba activa');
        els.chipStatus.text('En espera').removeClass('is-running is-finished');
        els.chipDate.text('');
        els.hint.html('Carga los equipos, elige quienes arrancan en <b>00:00:00</b> y el cronometro empieza a correr.');
      }
    }
    els.clockView.attr('class', 'clock ' + (test ? 'is-running' : 'is-idle'));
    els.clockMeta.text(test
      ? 'El cronometro corre desde el inicio de la prueba'
      : 'Esperando seleccion de equipos para iniciar');
    els.clockPulseText.text(test ? 'Corriendo' : 'En espera');
  }

  function renderPodium(test) {
    var entries = Core.testTeams(test).slice().sort(function (a, b) {
      var pa = a.position != null ? a.position : 9999;
      var pb = b.position != null ? b.position : 9999;
      if (pa !== pb) { return pa - pb; }
      var fa = a.finish === null ? Infinity : a.finish;
      var fb = b.finish === null ? Infinity : b.finish;
      if (fa !== fb) { return fa - fb; }
      return a.nombre.localeCompare(b.nombre, 'es');
    });

    var top = entries.slice(0, 3);
    var rest = entries.slice(3);

    var ties = Core.findTies(test);
    var pendingTies = ties.filter(function (g) { return !Core.tieResolved(test, g); }).length;
    var tieBtn = '';
    if (ties.length) {
      tieBtn = '<div style="margin-top:10px"><button class="btn btn--' + (pendingTies ? 'warn' : 'ghost') + ' btn--sm" data-act="tie-break">' +
        (pendingTies ? 'Resolver empates (' + pendingTies + ')' : 'Editar desempate') + '</button></div>';
    }

    var html = '<div class="podium__head">' +
      '<h2>Podio final</h2>' +
      '<p>' + Core.escape(test.nombre) + ' &middot; ' + Core.fmtDateTime(test.startedAt) +
      ' &middot; duracion total ' + Core.fmtFull(test.elapsedAtFinish || 0) + '</p>' +
      tieBtn +
      (Core.state.current ? '<div style="margin-top:10px"><button class="btn btn--ghost btn--sm" data-act="back-clock">Volver al cronometro</button></div>' : '') +
      '</div>';

    html += '<div class="podium__stage">';
    // izquierda = 2o, centro = 1o, derecha = 3o
    [1, 0, 2].forEach(function (idx, slot) {
      var e = top[idx];
      if (!e) { return; }
      var cls = slot === 0 ? 'podium__slot--2' : (slot === 1 ? 'podium__slot--1' : 'podium__slot--3');
      html += '<div class="podium__slot ' + cls + '">' +
        '<div class="podium__medal">' + (e.position != null ? e.position : '-') + '</div>' +
        '<div class="podium__name">' + Core.escape(e.nombre) + '</div>' +
        '<div class="podium__team">' + Core.escape(memberLabel(e) || 'sin integrantes') + '</div>' +
        '<div class="podium__time">' + Core.fmtFull(e.duration) + '</div>' +
        '<div class="podium__block"></div>' +
        '</div>';
    });
    html += '</div>';

    html += '<div class="podium__rest">';
    if (!rest.length) {
      html += '<div class="podium__title">Todos los equipos terminaron</div>';
    } else {
      html += '<div class="podium__title">Resto de equipos (' + rest.length + ')</div>';
      rest.forEach(function (e) {
        var dnf = e.finish === null;
        var pos = dnf ? '-' : e.position;
        html += '<div class="podium__restrow' + (dnf ? ' is-dnf' : '') + '">' +
          '<div class="podium__pos">' + pos + '</div>' +
          '<div><div class="podium__rname">' + Core.escape(e.nombre) + '</div>' +
          '<div class="podium__rteam">' + Core.escape(memberLabel(e) || 'sin integrantes') + '</div></div>' +
          '<div class="podium__rtime">' + (dnf ? (e.start === null ? 'no arranco' : 'no termino') : Core.fmtFull(e.duration)) + '</div>' +
          '</div>';
      });
    }
    html += '</div>';

    els.podiumView.html(html);
  }

  function renderPanel() {
    var q = (els.teamSearch.val() || '').toLowerCase().trim();
    var teams = Core.state.roster.teams;
    var html = '';
    var shown = 0;
    teams.forEach(function (t) {
      if (q && (t.nombre + ' ' + memberLabel(t)).toLowerCase().indexOf(q) === -1) { return; }
      shown++;
      html += '<li class="team">' +
        '<span class="team__dot"></span>' +
        '<span class="team__name" title="' + Core.escape(t.nombre) + '">' + Core.escape(t.nombre) + '</span>' +
        '<span class="team__members">' + Core.escape(memberLabel(t)) + '</span>' +
        '</li>';
    });
    els.teamList.html(html);
    els.teamCount.text(teams.length);
    els.teamEmpty.toggle(teams.length === 0);

    var test = Core.state.current;
    els.controlNote.text(test
      ? test.nombre + ' en curso - ' + Core.fmtFull(Core.elapsed())
      : 'Sin prueba activa. Pulsa "Nueva prueba" para cargar equipos y arrancar en 0.');

    var saveText = lastSaveAt
      ? 'Guardado automatico activo - ultimo guardado ' + lastSaveAt
      : 'Guardado automatico activo.';
    els.saveState.text(saveText);
  }

  function renderLive() {
    var test = Core.state.current;
    var body = els.liveBody;
    body.empty();
    if (!test) {
      els.liveEmpty.removeAttr('hidden');
      els.liveCount.text(0);
      return;
    }
    els.liveEmpty.attr('hidden', true);

    var entries = Core.testTeams(test);
    els.liveCount.text(entries.length);

    entries.forEach(function (e) {
      var st = Core.entryStatus(e);
      var badge = st === 'finished' ? '' :
        (st === 'running' ? '<span class="badge badge--live">en curso</span>' : '<span class="badge">sin iniciar</span>');
      var pos = e.position != null
        ? '<span class="pos pos--' + e.position + '">' + e.position + '</span>'
        : (st === 'running' ? '<span class="pos pos--live">-</span>' : '<span class="pos">-</span>');
      $('<tr></tr>')
        .attr('data-entry', e.id)
        .append(
          $('<td class="tdname"></td>').html(Core.escape(e.nombre) + ' ' + badge),
          $('<td class="mono dim"></td>').text(e.start === null ? '--' : Core.fmtFull(e.start)),
          $('<td class="mono" data-live="1"></td>').text(st === 'running' ? Core.fmtFull(Core.elapsed() - e.start) : '--'),
          $('<td class="mono"></td>').text(e.duration === null ? '--' : Core.fmtFull(e.duration)),
          $('<td class="num"></td>').html(pos)
        )
        .appendTo(body);
    });
  }

  function updateLiveTimes() {
    var test = Core.state.current;
    if (!test) { return; }
    var now = Core.elapsed();
    els.liveBody.find('tr').each(function () {
      var id = $(this).attr('data-entry');
      var e = test.entries[id];
      if (!e) { return; }
      var $cell = $(this).find('[data-live]');
      if (e.start !== null && e.finish === null) {
        $cell.text(Core.fmtFull(now - e.start));
      }
    });
  }

  function renderHistory() {
    var list = Core.state.history;
    els.histCount.text(list.length);
    els.historyList.empty();
    if (!list.length) {
      els.historyEmpty.removeAttr('hidden');
      return;
    }
    els.historyEmpty.attr('hidden', true);
    list.forEach(function (t) {
      var winners = Core.testTeams(t).filter(function (e) { return e.position != null; })
        .sort(function (a, b) { return a.position - b.position; })
        .slice(0, 3)
        .map(function (e) { return e.position + '. ' + Core.escape(e.nombre); }).join('  |  ');
      var $li = $('<li class="hist"></li>').append(
        $('<div class="hist__top"></div>').append(
          $('<span class="hist__name"></span>').text(t.nombre),
          $('<span class="hist__date"></span>').text(Core.fmtDateTime(t.startedAt))
        ),
        $('<div class="hist__podium"></div>').html(winners ? winners : 'sin equipos posicionados')
      );
      var $acts = $('<div class="hist__acts"></div>');
      $('<button class="btn btn--ghost btn--xs"></button>').text('Ver podio').on('click', function () {
        viewTest = t;
        renderStage();
        toast('Mostrando el podio de ' + t.nombre + '.', 'ok');
      }).appendTo($acts);
      $('<button class="btn btn--ghost btn--xs"></button>').text('Exportar').on('click', function () {
        doExportXLSX(t);
      }).appendTo($acts);
      $li.append($acts).appendTo(els.historyList);
    });
  }

  function renderButtons() {
    var running = !!Core.state.current;
    $('[data-act="start"]').prop('disabled', !running);
    $('[data-act="finish"]').prop('disabled', !running);
    $('[data-act="end"]').prop('disabled', !running);
  }

  /* =====================================================
     tick del reloj
     ===================================================== */
  function tick() {
    var ms = Core.elapsed();
    var hms = Core.fmtClock(ms);
    if (els.clockHms.text() !== hms) { els.clockHms.text(hms); }
    var cs = Core.fmtCs(ms);
    if (els.clockCs.text() !== cs) { els.clockCs.text(cs); }
    if (Core.state.current) {
      updateLiveTimes();
      els.controlNote.text(Core.state.current.nombre + ' en curso - ' + Core.fmtFull(ms));
    }
  }

  /* =====================================================
     exportaciones
     ===================================================== */
  function targetTest() {
    if (Core.state.current) { return Core.state.current; }
    if (viewTest) { return viewTest; }
    return Core.state.history[0] || null;
  }

  function doExportXLSX(test) {
    test = test || targetTest();
    if (!test) { toast('No hay pruebas para exportar.', 'warn'); return; }
    try {
      Storage.exportXLSX(test, Core.state.roster.teams);
      toast('Archivo Excel generado.', 'ok');
    } catch (e) {
      toast('No se pudo generar el Excel: ' + e.message, 'err');
    }
  }

  function doExportCSV(test) {
    test = test || targetTest();
    if (!test) { toast('No hay pruebas para exportar.', 'warn'); return; }
    Storage.exportResultsCSV(test);
    toast('Archivo CSV generado.', 'ok');
  }

  /* =====================================================
     init
     ===================================================== */
  function applyPanelVisibility() {
    if (!Core.state.ui) { Core.state.ui = { collapsed: [], panelHidden: false }; }
    var hidden = !!Core.state.ui.panelHidden;
    els.app.toggleClass('is-panel-hidden', hidden);
    els.btnPanel.attr('title', hidden ? 'Mostrar panel lateral' : 'Ocultar panel lateral')
      .attr('aria-label', hidden ? 'Mostrar panel lateral' : 'Ocultar panel lateral');
  }

  function togglePanel() {
    if (!Core.state.ui) { Core.state.ui = { collapsed: [], panelHidden: false }; }
    Core.state.ui.panelHidden = !Core.state.ui.panelHidden;
    applyPanelVisibility();
    save();
  }

  function bind() {
    els.app.on('click', '[data-act="toggle-panel"]', function () {
      togglePanel();
    });

    els.panel.on('click', '[data-act]', function () {
      var act = $(this).attr('data-act');
      switch (act) {
        case 'load-excel':
          pendingExcelAction = null;
          els.fileInput.val('').trigger('click');
          break;
        case 'clear-teams':
          if (!Core.state.roster.teams.length) { return; }
          openModal({
            title: 'Borrar equipos',
            sub: 'Se eliminan <b>' + Core.state.roster.teams.length + '</b> equipos de la lista.',
            showTools: false,
            html: '<p class="modal__msg">Esta accion no se puede deshacer.</p>',
            buttons: [
              { label: 'Cancelar', cls: 'btn btn--ghost', value: 'cancel' },
              { label: 'Borrar', cls: 'btn btn--danger', value: 'ok' }
            ],
            onConfirm: function () {
              Core.state.roster.teams = [];
              save();
              render();
              toast('Lista de equipos vaciada.', 'ok');
            }
          });
          break;
        case 'new-test':
          requestNewTest();
          break;
        case 'start':
          openStartPicker();
          break;
        case 'finish':
          openFinishPicker();
          break;
        case 'end':
          openFinishTest();
          break;
        case 'export-xlsx':
          doExportXLSX();
          break;
        case 'export-csv':
          doExportCSV();
          break;
        case 'export-roster':
          if (!Core.state.roster.teams.length) { toast('No hay equipos cargados.', 'warn'); break; }
          Storage.exportRosterCSV(Core.state.roster.teams);
          toast('Listado de equipos exportado.', 'ok');
          break;
        case 'export-json':
          Storage.exportJSON(Core.state);
          toast('Respaldo JSON generado.', 'ok');
          break;
        case 'import-json':
          els.jsonInput.val('').trigger('click');
          break;
        case 'wipe':
          openModal({
            title: 'Borrar todo',
            sub: 'Se eliminan equipos, prueba actual e historial de este navegador.',
            showTools: false,
            html: '<p class="modal__msg">Exporta un respaldo <b>.json</b> antes si quieres conservarlo.</p>',
            buttons: [
              { label: 'Cancelar', cls: 'btn btn--ghost', value: 'cancel' },
              { label: 'Borrar todo', cls: 'btn btn--danger', value: 'ok' }
            ],
            onConfirm: function () {
              Storage.clear();
              Core.state.roster.teams = [];
              Core.state.current = null;
              Core.state.history = [];
              Core.resetClock();
              viewTest = null;
              render();
              toast('Todo borrado.', 'ok');
            }
          });
          break;
      }
    });

    els.quick.on('click', '[data-act]', function () {
      var act = $(this).attr('data-act');
      if (act === 'new-test') { requestNewTest(); }
      else if (act === 'start') { openStartPicker(); }
      else if (act === 'finish') { openFinishPicker(); }
      else if (act === 'end') { openFinishTest(); }
    });

    els.fileInput.on('change', function () {
      var f = this.files && this.files[0];
      if (f) { handleExcelFile(f); }
      this.value = '';
    });

    els.jsonInput.on('change', function () {
      var f = this.files && this.files[0];
      if (!f) { return; }
      Storage.importJSON(f).then(function (data) {
        Core.state.roster = data.roster || { teams: [] };
        Core.state.history = data.history || [];
        Core.state.current = data.current || null;
        Core.state.ui = data.ui || { collapsed: [] };
        if (!Array.isArray(Core.state.ui.collapsed)) { Core.state.ui.collapsed = []; }
        applyCollapsed();
        applyPanelVisibility();
        if (Core.state.current) {
          Core.state.current.status = 'running';
          Core.resetClock();
          if (typeof data.clockElapsed === 'number' && data.clockElapsed > 0) {
            Core.state.clock.base = performance.now() - data.clockElapsed;
            Core.state.clock.running = true;
          } else {
            Core.startClock();
          }
          viewTest = null;
        }
        save();
        render();
        toast('Respaldo importado.', 'ok');
      }).catch(function (err) {
        toast(err.message, 'err');
      });
      this.value = '';
    });

    els.teamSearch.on('input', renderPanel);

    els.modal.on('click', '[data-close]', closeModal);
    els.podiumView.on('click', '[data-act="back-clock"]', function () {
      viewTest = null;
      render();
    });
    els.podiumView.on('click', '[data-act="tie-break"]', function () {
      openTieBreak(viewTest);
    });
    $(document).on('keydown', function (e) {
      if (e.key === 'Escape' && !els.modal.attr('hidden')) { closeModal(); }
    });

    // drag & drop de archivos
    ['dragenter', 'dragover'].forEach(function (ev) {
      els.dropzone.on(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        els.dropzone.addClass('is-over');
      });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      els.dropzone.on(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        els.dropzone.removeClass('is-over');
      });
    });
    els.dropzone.on('drop', function (e) {
      var f = e.originalEvent && e.originalEvent.dataTransfer && e.originalEvent.dataTransfer.files[0];
      if (f) { pendingExcelAction = null; handleExcelFile(f); }
    });

    $(window).on('beforeunload', function () { save(); });
    document.addEventListener('visibilitychange', function () {
      if (!document.hidden) { tick(); renderLive(); }
    });
  }

  function restore() {
    var data = Storage.load();
    if (!data) { return false; }
    Core.state.roster = data.roster || { teams: [] };
    Core.state.history = data.history || [];
    Core.state.current = data.current || null;
    Core.state.ui = data.ui || { collapsed: [] };
    if (!Array.isArray(Core.state.ui.collapsed)) { Core.state.ui.collapsed = []; }
    if (Core.state.current && Core.state.current.status === 'running') {
      // la prueba sigue: el reloj se reanuda desde el tiempo guardado
      Core.resetClock();
      if (typeof data.clockElapsed === 'number' && data.clockElapsed > 0) {
        Core.state.clock.base = performance.now() - data.clockElapsed;
        Core.state.clock.running = true;
      } else {
        Core.startClock();
      }
    }
    return true;
  }

  $(function () {
    els = {
      app: $('.app'),
      btnPanel: $('#btnPanel'),
      panel: $('#panel'),
      quick: $('#quickActions'),
      testName: $('#testName'),
      chipStatus: $('#chipStatus'),
      chipDate: $('#chipClockDate'),
      clockView: $('#clockView'),
      clockHms: $('#clockHms'),
      clockCs: $('#clockCs'),
      clockMeta: $('#clockMeta'),
      clockPulseText: $('#clockPulseText'),
      podiumView: $('#podiumView'),
      hint: $('#hint'),
      teamList: $('#teamList'),
      teamEmpty: $('#teamEmpty'),
      teamCount: $('#teamCount'),
      teamSearch: $('#teamSearch'),
      dropzone: $('#dropzone'),
      fileInput: $('#fileInput'),
      jsonInput: $('#jsonInput'),
      liveBody: $('#liveBody'),
      liveEmpty: $('#liveEmpty'),
      liveCount: $('#liveCount'),
      controlNote: $('#controlNote'),
      historyList: $('#historyList'),
      historyEmpty: $('#historyEmpty'),
      histCount: $('#histCount'),
      saveState: $('#saveState'),
      modal: $('#modal'),
      modalTitle: $('#modalTitle'),
      modalSub: $('#modalSub'),
      modalBody: $('#modalBody'),
      modalFoot: $('#modalFoot'),
      modalTools: $('#modalTools'),
      modalSearch: $('#modalSearch'),
      modalAll: $('#modalAll'),
      modalNone: $('#modalNone'),
      toasts: $('#toasts'),
      btnStart: $('[data-act="start"]').first(),
      btnFinish: $('[data-act="finish"]').first(),
      btnEnd: $('[data-act="end"]').first()
    };

    bind();
    initCollapsible();
    var hadState = restore();
    applyCollapsed();
    applyPanelVisibility();
    render();
    setInterval(tick, 50);
    tick();

    if (Core.state.current) {
      toast('Prueba "' + Core.state.current.nombre + '" recuperada. El cronometro sigue corriendo.', 'ok', 5000);
    } else if (!hadState || !Core.state.history.length) {
      setTimeout(function () { openSetup(); }, 350);
    }
  });

})(jQuery);
