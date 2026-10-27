/* A tela: botões, filtros, abas, tabelas e exportação. As contas ficam em regras.js. */
(function () {
  'use strict';
  var R = window.Regras, T = window.TrelloApi, X = window.ExcelFlow;
  var t = TrelloPowerUp.iframe({ appKey: window.APP_KEY, appName: window.APP_NAME });
  var $ = function (id) { return document.getElementById(id); };
  var esc = R.esc;

  var PALETA = ['#4C78A8', '#F58518', '#54A24B', '#E45756', '#B279A2', '#72B7B2', '#EECA3B', '#9D755D', '#FF9DA6', '#BAB0AC'];
  var COR_ETIQUETA = { green: '#4bce97', yellow: '#f5cd47', orange: '#fea362', red: '#f87168', purple: '#9f8fef',
    blue: '#579dff', sky: '#6cc3e0', lime: '#94c748', pink: '#e774bb', black: '#8590a2' };
  function corEtiqueta(c) { return COR_ETIQUETA[String(c || '').split('_')[0]] || '#8590a2'; }

  var ICONES = {
    atualizar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><path d="M21 3v6h-6"/></svg>',
    fechar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    entrada: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><path d="M10 17l5-5-5-5M15 12H3"/></svg>',
    saida: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5M21 12H9"/></svg>',
    cancelar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M15 9l-6 6M9 9l6 6"/></svg>',
    saldo: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v8M8 8h8M8 19h8"/></svg>'
  };

  var S = {
    boardId: null, token: null, boardNome: '', listas: [], labels: [],
    prefs: { periodo: '30', de: '', ate: '', etiqueta: '', aba: 'resumo', abertos: {} },
    cfg: { listaDesist: '', janela: true, dedup: true },
    cache: null, reqId: 0, res: null, tabelas: {}, avisoCarga: '', decisoes: {}, filtroTipo: ''
  };

  // ---------- Tema ----------
  function aplicarTema() {
    var th = null;
    try { th = t.getContext().theme; } catch (e) {}
    if (!th) th = (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
    document.body.className = th === 'dark' ? 'dark' : 'light';
  }

  // ---------- Período ----------
  function intervalo() {
    var p = S.prefs.periodo, now = new Date(), since, until = null, key = p;
    if (p === 'custom') {
      if (!S.prefs.de || !S.prefs.ate) throw new Error('Escolha a data inicial e a data final.');
      since = new Date(S.prefs.de + 'T00:00:00'); until = new Date(S.prefs.ate + 'T23:59:59.999');
      if (since > until) throw new Error('A data inicial é depois da data final.');
      key = 'custom:' + S.prefs.de + ':' + S.prefs.ate;
    } else if (p === 'mes') {
      since = new Date(now.getFullYear(), now.getMonth(), 1);
    } else if (p === 'mespassado') {
      since = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      until = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999);
    } else {
      since = new Date(now.getTime() - parseInt(p, 10) * 86400000);
    }
    var fim = until || now, dur = fim - since, span = dur / 86400000;
    return {
      since: since, until: until, fim: fim, span: span, key: key,
      comparar: span <= 190, prevSince: new Date(since.getTime() - dur),
      texto: R.fmtFull(since) + ' a ' + R.fmtFull(fim)
    };
  }

  // ---------- Carregamento ----------
  function setCarregando(sim, msg) {
    $('principal').classList.toggle('carregando', !!sim);
    ['btnAtualizar', 'btnExportar'].forEach(function (id) { $(id).disabled = !!sim; });
    $('status').textContent = sim ? (msg || 'Carregando...') : '';
  }
  function mostrarAviso(html, tipo) {
    $('aviso').innerHTML = html ? '<div class="aviso ' + (tipo || 'atencao') + '">' + html + '</div>' : '';
  }
  function mostrarAvisos(lista) {
    $('aviso').innerHTML = lista.filter(function (a) { return a && a.html; })
      .map(function (a) { return '<div class="aviso ' + (a.tipo || 'atencao') + '">' + a.html + '</div>'; }).join('');
  }
  function toast(msg, acao) {
    var antigo = document.querySelector('.toast'); if (antigo) antigo.remove();
    var el = document.createElement('div'); el.className = 'toast'; el.setAttribute('role', 'status');
    el.innerHTML = '<span>' + esc(msg) + '</span>' + (acao ? '<button type="button">' + esc(acao.txt) + '</button>' : '');
    if (acao) el.querySelector('button').onclick = function () { el.remove(); acao.fn(); };
    document.body.appendChild(el);
    setTimeout(function () { if (el.parentNode) el.remove(); }, acao ? 12000 : 5000);
  }

  async function carregar(force) {
    var id = ++S.reqId;
    setCarregando(true);
    mostrarAviso('');
    try {
      var iv = intervalo(), c = S.cache;
      var fresco = c && c.key === iv.key && !force && (Date.now() - c.quando < 600000);
      if (!fresco) {
        var desde = iv.comparar ? iv.prevSince : iv.since;
        setCarregando(true, 'Buscando as movimentações no Trello...');
        var actions = await T.fetchActions(S.boardId, desde, iv.until, S.token);
        if (id !== S.reqId) return;
        var cardLabels = {}, cardInfo = {};
        S.avisoCarga = '';
        try { var cs = await T.fetchCards(S.boardId, S.token); cardLabels = cs.labels; cardInfo = cs.info; }
        catch (e) { S.avisoCarga = 'Não consegui ler as etiquetas e descrições dos cards (' + e.message + '). Os números gerais estão corretos, mas a leitura de 2 projetos e as etiquetas podem faltar.'; }
        if (id !== S.reqId) return;
        var nomes = {};
        await T.resolverBoards(actions, S.token, nomes);
        if (id !== S.reqId) return;
        S.cache = { key: iv.key, quando: Date.now(), iv: iv, actions: actions, cardLabels: cardLabels, cardInfo: cardInfo, nomes: nomes, hist: null };
      }
      await calcular(id);
    } catch (e) {
      if (id === S.reqId && (e.status === 401 || e.status === 403)) {
        $('app').hidden = true; $('auth').hidden = false; // autorização vencida ou revogada: pede de novo
      } else if (id === S.reqId) {
        mostrarAviso(esc(e.message) + ' <button class="btn" id="btnTentar">Tentar de novo</button>', 'erro');
        var b = $('btnTentar'); if (b) b.onclick = function () { carregar(true); };
      }
    } finally {
      if (id === S.reqId) setCarregando(false);
    }
  }

  // Recalcula com os dados já buscados (mudou filtro, sem nova consulta)
  async function recalcular() {
    if (!S.cache) return carregar(false);
    var id = ++S.reqId;
    setCarregando(true);
    try { await calcular(id); }
    catch (e) { if (id === S.reqId) mostrarAviso(esc(e.message), 'erro'); }
    finally { if (id === S.reqId) setCarregando(false); }
  }

  async function garantirHistorico(clientes, id) {
    var c = S.cache, minRef = null, maxFim = null;
    clientes.forEach(function (cl) {
      if (!minRef || cl.entrada < minRef) minRef = cl.entrada;
      if (!maxFim || cl.saida > maxFim) maxFim = cl.saida;
    });
    var ini = new Date(minRef.getTime() - 60000), fim = new Date(maxFim.getTime() + 1000), h = c.hist;
    if (!h || ini < h.ini || fim > h.fim) {
      var uIni = h && h.ini < ini ? h.ini : ini, uFim = h && h.fim > fim ? h.fim : fim;
      setCarregando(true, 'Calculando o tempo em cada coluna...');
      var acts = await T.fetchActions(S.boardId, uIni, uFim, S.token, T.HIST_FILTER);
      if (id !== S.reqId) return false;
      var porCard = {};
      acts.forEach(function (a) {
        var cid = a.data && a.data.card && a.data.card.id;
        if (!cid) return;
        (porCard[cid] = porCard[cid] || []).push({ type: a.type, date: new Date(a.date), data: a.data, id: a.id });
      });
      c.hist = { ini: uIni, fim: uFim, porCard: porCard };
    }
    R.aplicarColunas(clientes, c.hist.porCard);
    return true;
  }

  async function calcular(id) {
    var c = S.cache, iv = c.iv;
    var cfg = { listaDesist: S.cfg.listaDesist, boardNomes: c.nomes };
    var evs = R.montarEventos(c.actions, cfg, S.prefs.etiqueta, c.cardLabels);
    var nomeEtiq = {};
    S.labels.forEach(function (l) { nomeEtiq[l.id] = l.name || ''; });
    evs.forEach(function (ev) {
      var info = c.cardInfo && c.cardInfo[ev.id];
      var etNomes = ((c.cardLabels && c.cardLabels[ev.id]) || []).map(function (lid) { return nomeEtiq[lid] || ''; });
      ev.doisProjetos = R.leDoisProjetos([ev.card.name, info && info.name], etNomes, [info && info.desc]);
    });
    evs.forEach(function (ev) {
      if (!ev.c.copia) return;
      var d = ev.a.data || {}, src = d.cardSource, bs = d.boardSource;
      // cópia "interna" = o card de origem é deste mesmo board (vendas copiando de outro board é o normal)
      if (bs && bs.id) ev.c.copiaInterna = bs.id === S.boardId;
      else ev.c.copiaInterna = !!(src && src.id && c.cardLabels && (src.id in c.cardLabels));
    });
    var kept = R.processar(evs, S.cfg.janela ? 5 * 60000 : 0);
    R.aplicarDecisoes(kept, S.decisoes);
    if (S.cfg.dedup) R.deduplicarNumeros(kept);
    R.marcarPendencias(kept);

    var noAtual = function (ev) { return ev.when >= iv.since && (!iv.until || ev.when <= iv.until); };
    var atual = kept.filter(noAtual);
    var anterior = iv.comparar ? kept.filter(function (ev) { return ev.when >= iv.prevSince && ev.when < iv.since; }) : [];
    var contados = atual.filter(function (e) { return !e.ign; });
    var ignorados = atual.filter(function (e) { return e.ign; });
    var pendentes = contados.filter(function (e) { return e.pend && e.pend.length; });
    var decididos = atual.filter(function (e) { return e.decidido; });

    var arq = {}, exc = {};
    evs.forEach(function (ev) {
      if (ev.tipo !== 'rem' || !noAtual(ev)) return;
      (ev.c.sub === 'excluido' ? exc : arq)[ev.id] = 1;
    });

    var clientes = R.montarClientes(contados, c.cardLabels);
    var avisoCol = '';
    if (clientes.length) {
      try {
        var ok = await garantirHistorico(clientes, id);
        if (ok === false) return;
      } catch (e) { avisoCol = 'Não consegui calcular o tempo por coluna: ' + e.message; }
    }
    if (id !== S.reqId) return;
    var colunas = R.colunasUsadas(clientes, S.listas);
    var kpis = R.contar(contados);
    kpis.prev = iv.comparar ? R.contar(anterior) : null;

    var linhas = contados.map(function (ev) {
      return { date: ev.when, tipo: ev.tipo, motivo: ev.c.motivo, card: ev.card.name || '(sem nome)', link: urlCard(ev.card),
               origem: ev.c.origem, destino: ev.c.destino,
               quem: (ev.a.memberCreator && ev.a.memberCreator.fullName) || '',
               obs: ev.decidido ? 'Confirmado por você: ' + (ev.decidido === 'c' ? 'conta' : 'não conta')
                  : (ev.pend && ev.pend.length ? 'Pendente: ' + ev.pend.join(' · ') : '') };
    }).sort(function (a, b) { return b.date - a.date; });

    S.res = {
      iv: iv, kpis: kpis, serieRotulo: iv.span > 31 ? 'Semana de' : 'Dia', serie: R.serie(contados, iv.span),
      etiquetas: R.porEtiqueta(contados, clientes, c.cardLabels, S.labels),
      destinos: R.destinos(contados), etapas: R.etapasDesistencia(contados),
      colunas: colunas, clientes: clientes, gargalos: R.gargalos(clientes, colunas),
      linhas: linhas, ignorados: ignorados, pendentes: pendentes, decididos: decididos,
      naoContados: { arquivados: Object.keys(arq).length, excluidos: Object.keys(exc).length }
    };
    mostrarAvisos([
      avisoCol ? { html: esc(avisoCol), tipo: 'erro' } : (S.avisoCarga ? { html: esc(S.avisoCarga) } : null),
      S.cfg.listaDesist ? null : { html: 'Escolha em <b>Filtros</b> a coluna de desistentes para contar os cancelamentos. <button class="btn" id="btnAbrirFiltros">Abrir filtros</button>' },
      pendentes.length ? { html: '<span><b>' + pendentes.length + ' pendente' + (pendentes.length === 1 ? '' : 's') + ' de confirmação.</b> Já estão contados provisoriamente. <button class="btn" id="btnVerPendentes">Revisar</button></span>', tipo: 'info' } : null
    ]);
    var bf = $('btnAbrirFiltros'); if (bf) bf.onclick = function () { alternarFiltros(true); };
    var bp = $('btnVerPendentes'); if (bp) bp.onclick = function () { mostrarAba('pendentes'); };
    renderTudo();
  }

  // ---------- Tabelas (ordenar, buscar, mostrar mais) ----------
  // Abre o card em outra aba do navegador (cards excluídos não têm mais página)
  function urlCard(card, excluido) {
    var ref = card && (card.shortLink || card.id);
    return !excluido && ref ? 'https://trello.com/c/' + ref : '';
  }
  function linkCard(url, html) {
    return url ? '<a class="lk" href="' + esc(url) + '" target="_blank" rel="noopener noreferrer" title="Abrir o card no Trello">' + html + '</a>' : html;
  }
  function linkEv(ev, html) { return linkCard(urlCard(ev.card, ev.c && ev.c.sub === 'excluido'), html); }

  function criarTabela(el, cols, o) {
    o = o || {};
    var st = { rows: [], sortK: o.sortK || null, dir: o.dir || -1, limite: 50, busca: '' };
    function lista() {
      var r = st.rows;
      if (st.extra) r = r.filter(st.extra);
      if (st.busca) {
        var q = st.busca.toLowerCase();
        r = r.filter(function (x) { return o.texto(x).toLowerCase().indexOf(q) >= 0; });
      }
      var col = cols.filter(function (c) { return c.k === st.sortK; })[0];
      if (col) {
        r = r.slice().sort(function (a, b) {
          var va = col.v(a), vb = col.v(b);
          return (va < vb ? -1 : va > vb ? 1 : 0) * st.dir;
        });
      }
      return r;
    }
    function render() {
      var r = lista(), total = r.length;
      if (o.contagem) o.contagem(total, st.rows.length);
      if (!total) {
        el.innerHTML = '<div class="vazio"><b>' + esc(st.rows.length ? 'Nada encontrado' : (o.vazioTitulo || 'Sem dados')) + '</b>' +
          esc(st.rows.length ? 'Tente outra busca.' : (o.vazioTexto || 'Nada para mostrar neste período.')) + '</div>';
        return;
      }
      var h = '<table><thead><tr>' + cols.map(function (c) {
        var ativo = c.k === st.sortK, seta = ativo ? (st.dir > 0 ? '▲' : '▼') : '↕';
        return '<th class="' + (c.cls || '') + (c.v && !o.semOrdem ? ' ord' : '') + (ativo ? ' ativo' : '') + '" data-k="' + c.k + '"' +
          (c.v && !o.semOrdem ? ' tabindex="0" role="button" aria-label="Ordenar por ' + esc(c.t) + '"' : '') + '>' + esc(c.t) +
          (c.v && !o.semOrdem ? '<span class="seta">' + seta + '</span>' : '') + '</th>';
      }).join('') + '</tr></thead><tbody>';
      r.slice(0, st.limite).forEach(function (x) {
        h += '<tr>' + cols.map(function (c) { return '<td class="' + (c.cls || '') + '">' + c.h(x) + '</td>'; }).join('') + '</tr>';
      });
      h += '</tbody></table>';
      if (total > st.limite) h += '<button class="btn mais" data-mais>Mostrar mais (' + (total - st.limite) + ' restantes)</button>';
      el.innerHTML = h;
    }
    function ordenar(k) {
      var col = cols.filter(function (c) { return c.k === k; })[0];
      if (!col || !col.v || o.semOrdem) return;
      if (st.sortK === k) st.dir = -st.dir; else { st.sortK = k; st.dir = col.cls === 'num' ? -1 : 1; }
      render();
    }
    el.addEventListener('click', function (e) {
      if (e.target.closest('[data-mais]')) { st.limite += 50; render(); return; }
      var th = e.target.closest('th[data-k]'); if (th) ordenar(th.getAttribute('data-k'));
    });
    el.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      var th = e.target.closest('th[data-k]'); if (th) { e.preventDefault(); ordenar(th.getAttribute('data-k')); }
    });
    return {
      set: function (rows) { st.rows = rows; st.limite = 50; render(); },
      busca: function (q) { st.busca = q || ''; st.limite = 50; render(); },
      filtrar: function (fn) { st.extra = fn || null; st.limite = 50; render(); }
    };
  }

  function celBarra(valor, max, cor, texto) {
    var pct = max > 0 ? Math.max(0, valor) / max * 100 : 0;
    return '<div class="cel-barra"><span class="fundo" style="width:' + pct.toFixed(1) + '%;background:' + cor + '"></span><span class="v">' + texto + '</span></div>';
  }
  function saldoHtml(v) { return '<b style="color:var(--' + (v < 0 ? 'cancel' : v > 0 ? 'in' : 'muted') + ')">' + (v > 0 ? '+' : '') + v + '</b>'; }
  function taxaTxt(v) { return v == null ? '—' : (Math.round(v * 1000) / 10).toString().replace('.', ',') + '%'; }
  var PILL = { in: ['in', 'Entrada'], out: ['out', 'Saída'], cancel: ['cancel', 'Cancelamento'] };
  function pill(tipo) { return '<span class="pill ' + PILL[tipo][0] + '">' + PILL[tipo][1] + '</span>'; }

  function criarTabelas() {
    if (S.tabelas.serie) return; // já criadas (evita ligar os cliques duas vezes)
    S.tabelas.serie = criarTabela($('tSerie'), [
      { k: 'k', t: 'Período', v: function (r) { return r.k; }, h: function (r) { return '<b>' + (S.res.serieRotulo === 'Semana de' ? 'Semana de ' : '') + R.fmtFull(new Date(r.k)) + '</b>'; } },
      { k: 'in', t: 'Entradas', cls: 'num', v: function (r) { return r.in; }, h: function (r) { return celBarra(r.in, S.max.in, 'var(--in)', r.in); } },
      { k: 'out', t: 'Saídas', cls: 'num', v: function (r) { return r.out; }, h: function (r) { return celBarra(r.out, S.max.out, 'var(--out)', r.out); } },
      { k: 'cancel', t: 'Cancelamentos', cls: 'num', v: function (r) { return r.cancel; }, h: function (r) { return celBarra(r.cancel, S.max.cancel, 'var(--cancel)', r.cancel); } },
      { k: 'saldo', t: 'Saldo', cls: 'num', v: function (r) { return r.saldo; }, h: function (r) { return saldoHtml(r.saldo); } }
    ], { sortK: 'k', dir: -1, vazioTitulo: 'Nenhuma movimentação', vazioTexto: 'Não houve entradas, saídas nem cancelamentos neste período.' });

    S.tabelas.etiquetas = criarTabela($('tEtiquetas'), [
      { k: 'nome', t: 'Etiqueta', v: function (r) { return r.nome.toLowerCase(); },
        h: function (r) { return '<span class="ponto" style="background:' + (r.cor ? corEtiqueta(r.cor) : 'var(--line)') + '"></span><b>' + esc(r.nome) + '</b>'; } },
      { k: 'in', t: 'Entradas', cls: 'num', v: function (r) { return r.in; }, h: function (r) { return r.in; } },
      { k: 'out', t: 'Saídas', cls: 'num', v: function (r) { return r.out; }, h: function (r) { return r.out; } },
      { k: 'cancel', t: 'Cancelamentos', cls: 'num', v: function (r) { return r.cancel; }, h: function (r) { return r.cancel; } },
      { k: 'taxa', t: 'Taxa de cancel.', cls: 'num', v: function (r) { return r.taxa == null ? -1 : r.taxa; }, h: function (r) { return taxaTxt(r.taxa); } },
      { k: 'saldo', t: 'Saldo', cls: 'num', v: function (r) { return r.saldo; }, h: function (r) { return saldoHtml(r.saldo); } },
      { k: 'tempo', t: 'Tempo médio', cls: 'num', v: function (r) { return r.tempoMedio == null ? -1 : r.tempoMedio; },
        h: function (r) { return r.tempoMedio == null ? '—' : R.fmtTempo(r.tempoMedio); } }
    ], { sortK: null, vazioTitulo: 'Sem dados por etiqueta', vazioTexto: 'Nenhum card com movimentação neste período.' });

    S.tabelas.destinos = criarTabela($('tDestinos'), [
      { k: 'nome', t: 'Board', v: function (r) { return r.nome.toLowerCase(); }, h: function (r) { return esc(r.nome); } },
      { k: 'qtd', t: 'Clientes', cls: 'num', v: function (r) { return r.qtd; }, h: function (r) { return celBarra(r.qtd, S.max.destino, 'var(--out)', r.qtd); } }
    ], { sortK: 'qtd', dir: -1, vazioTitulo: 'Nenhuma saída', vazioTexto: 'Nenhum cliente foi enviado a outro board.' });

    S.tabelas.etapas = criarTabela($('tEtapas'), [
      { k: 'nome', t: 'Etapa (coluna)', v: function (r) { return r.nome.toLowerCase(); }, h: function (r) { return esc(r.nome); } },
      { k: 'qtd', t: 'Cancelamentos', cls: 'num', v: function (r) { return r.qtd; }, h: function (r) { return celBarra(r.qtd, S.max.etapa, 'var(--cancel)', r.qtd); } }
    ], { sortK: 'qtd', dir: -1, vazioTitulo: 'Nenhum cancelamento', vazioTexto: 'Nenhum cliente foi para desistentes.' });

    S.tabelas.gargalos = criarTabela($('tGargalos'), [
      { k: 'nome', t: 'Coluna', v: function (r) { return r.nome.toLowerCase(); }, h: function (r) {
          var i = S.res.colunas.map(function (c) { return c.id; }).indexOf(r.id);
          return '<span class="ponto" style="background:' + PALETA[Math.max(i, 0) % PALETA.length] + '"></span><b>' + esc(r.nome) + '</b>'; } },
      { k: 'n', t: 'Clientes', cls: 'num', v: function (r) { return r.n; }, h: function (r) { return r.n; } },
      { k: 'media', t: 'Tempo médio', cls: 'num', v: function (r) { return r.media; }, h: function (r) { return celBarra(r.media, S.max.media, 'var(--warn)', '<b>' + R.fmtTempo(r.media) + '</b>'); } },
      { k: 'mediana', t: 'Mediana', cls: 'num', v: function (r) { return r.mediana; }, h: function (r) { return R.fmtTempo(r.mediana); } },
      { k: 'maior', t: 'Maior tempo', cls: 'num', v: function (r) { return r.maior; }, h: function (r) { return R.fmtTempo(r.maior); } },
      { k: 'share', t: '% do tempo total', cls: 'num', v: function (r) { return r.share; }, h: function (r) { return taxaTxt(r.share); } }
    ], { sortK: 'media', dir: -1, vazioTitulo: 'Sem dados de tempo', vazioTexto: 'Nenhum cliente saiu ou cancelou neste período.' });

    S.tabelas.clientes = criarTabela($('tClientes'), [
      { k: 'numero', t: 'Nº', v: function (r) { return r.numero; }, h: function (r) { return r.numero ? linkCard(urlCard({ shortLink: r.shortLink, id: r.id }), esc(r.numero)) : '—'; } },
      { k: 'nome', t: 'Nome', v: function (r) { return (r.nome || r.card).toLowerCase(); }, h: function (r) { return linkCard(urlCard({ shortLink: r.shortLink, id: r.id }), '<b>' + esc(r.nome || r.card) + '</b>'); } },
      { k: 'situacao', t: 'Situação', v: function (r) { return r.situacao; }, h: function (r) {
          return r.situacao === 'out' ? '<span class="pill out">Saiu → ' + esc(r.destino) + '</span>' : '<span class="pill cancel">Cancelou' + (r.etapa ? ' em ' + esc(r.etapa) : '') + '</span>'; } },
      { k: 'entrada', t: 'Entrada', v: function (r) { return r.entrada.getTime(); }, h: function (r) { return R.fmtFull(r.entrada); } },
      { k: 'saida', t: 'Saída', v: function (r) { return r.saida.getTime(); }, h: function (r) { return R.fmtFull(r.saida); } },
      { k: 'ms', t: 'Tempo', cls: 'num', v: function (r) { return r.ms; }, h: function (r) { return '<b>' + R.fmtTempo(r.ms) + '</b>'; } },
      { k: 'colunas', t: 'Tempo por coluna', v: null, h: barraColunas }
    ], { sortK: 'saida', dir: -1, vazioTitulo: 'Nenhum cliente', vazioTexto: 'Nenhum cliente saiu para outro board ou cancelou neste período.',
         texto: function (r) { return r.numero + ' ' + r.nome + ' ' + r.card; },
         contagem: function (n, total) { $('contaClientes').textContent = total ? n + ' de ' + total + ' clientes' : ''; } });

    S.tabelas.cards = criarTabela($('tCards'), [
      { k: 'date', t: 'Data', v: function (r) { return r.date.getTime(); }, h: function (r) { return R.fmtFull(r.date) + ' <span class="muted">' + R.fmtHora(r.date) + '</span>'; } },
      { k: 'tipo', t: 'Tipo', v: function (r) { return r.tipo; }, h: function (r) { return pill(r.tipo); } },
      { k: 'card', t: 'Card', v: function (r) { return r.card.toLowerCase(); }, h: function (r) { return linkCard(r.link, '<b>' + esc(r.card) + '</b>'); } },
      { k: 'motivo', t: 'Motivo', v: function (r) { return r.motivo; }, h: function (r) { return esc(r.motivo); } },
      { k: 'origem', t: 'Lista de origem', v: function (r) { return r.origem; }, h: function (r) { return esc(r.origem) || '—'; } },
      { k: 'destino', t: 'Lista de destino', v: function (r) { return r.destino; }, h: function (r) { return esc(r.destino) || '—'; } },
      { k: 'quem', t: 'Responsável', v: function (r) { return r.quem; }, h: function (r) { return esc(r.quem) || '—'; } }
    ], { sortK: 'date', dir: -1, vazioTitulo: 'Nenhum card', vazioTexto: 'Nenhuma movimentação contada neste período.',
         texto: function (r) { return r.card + ' ' + r.motivo + ' ' + r.origem + ' ' + r.destino + ' ' + r.quem; },
         contagem: function (n, total) { $('contaCards').textContent = total ? n + ' de ' + total + ' movimentações' : ''; } });

    S.tabelas.ignorados = criarTabela($('tIgnorados'), [
      { k: 'date', t: 'Data', v: function (r) { return r.when.getTime(); }, h: function (r) { return R.fmtFull(r.when) + ' <span class="muted">' + R.fmtHora(r.when) + '</span>'; } },
      { k: 'tipo', t: 'Evento', v: function (r) { return r.tipo; }, h: function (r) { return pill(r.tipo); } },
      { k: 'card', t: 'Card', v: function (r) { return (r.card.name || '').toLowerCase(); }, h: function (r) { return linkEv(r, '<b>' + esc(r.card.name || '(sem nome)') + '</b>'); } },
      { k: 'motivo', t: 'O que foi', v: function (r) { return r.c.motivo; }, h: function (r) { return esc(r.c.motivo); } },
      { k: 'porque', t: 'Por que foi ignorado', v: function (r) { return r.ign; }, h: function (r) { return esc(r.ign); } },
      { k: 'acao', t: 'Ação', v: null, h: function (r) {
          return r.decidido === 'n' ? botaoDecisao(r, 'x', 'Desfazer') : botaoDecisao(r, 'c', 'Contar mesmo assim'); } }
    ], { sortK: 'date', dir: -1, vazioTitulo: 'Nada ignorado', vazioTexto: 'Nenhum evento foi ignorado neste período.' });

    S.tabelas.pendentes = criarTabela($('tPendentes'), [
      { k: 'date', t: 'Data', v: function (r) { return r.when.getTime(); }, h: function (r) { return R.fmtFull(r.when) + ' <span class="muted">' + R.fmtHora(r.when) + '</span>'; } },
      { k: 'tipo', t: 'Evento', v: function (r) { return r.tipo; }, h: function (r) { return pill(r.tipo); } },
      { k: 'card', t: 'Card', v: function (r) { return (r.card.name || '').toLowerCase(); }, h: function (r) { return linkEv(r, '<b>' + esc(r.card.name || '(sem nome)') + '</b>'); } },
      { k: 'pend', t: 'Por que precisa de confirmação', v: function (r) { return r.pend.join(' '); }, h: function (r) { return esc(r.pend.join(' · ')); } },
      { k: 'acao', t: 'Contar?', v: null, h: function (r) { return botaoDecisao(r, 'c', 'Conta') + ' ' + botaoDecisao(r, 'n', 'Não conta'); } }
    ], { sortK: 'date', dir: -1, vazioTitulo: 'Nada pendente', vazioTexto: 'Nenhum caso precisa da sua confirmação neste período.',
         texto: function (r) { return (r.card.name || '') + ' ' + r.pend.join(' '); } });

    S.tabelas.decididos = criarTabela($('tDecididos'), [
      { k: 'date', t: 'Data', v: function (r) { return r.when.getTime(); }, h: function (r) { return R.fmtFull(r.when) + ' <span class="muted">' + R.fmtHora(r.when) + '</span>'; } },
      { k: 'tipo', t: 'Evento', v: function (r) { return r.tipo; }, h: function (r) { return pill(r.tipo); } },
      { k: 'card', t: 'Card', v: function (r) { return (r.card.name || '').toLowerCase(); }, h: function (r) { return linkEv(r, '<b>' + esc(r.card.name || '(sem nome)') + '</b>'); } },
      { k: 'dec', t: 'Sua decisão', v: function (r) { return r.decidido; }, h: function (r) { return r.decidido === 'c' ? '<b style="color:var(--in)">Conta</b>' : '<b style="color:var(--cancel)">Não conta</b>'; } },
      { k: 'acao', t: 'Ação', v: null, h: function (r) { return botaoDecisao(r, 'x', 'Desfazer'); } }
    ], { sortK: 'date', dir: -1, vazioTitulo: 'Nenhuma decisão', vazioTexto: 'Você ainda não confirmou nenhum caso neste período.' });
  }

  function botaoDecisao(ev, dec, rotulo) {
    return '<button class="btn mini-btn' + (dec === 'c' ? ' ok' : dec === 'n' ? ' nao' : '') + '" data-dec="' + dec + '" data-id="' + esc(ev.a.id) + '">' + esc(rotulo) + '</button>';
  }

  function barraColunas(c) {
    var total = 0, partes = [];
    S.res.colunas.forEach(function (col, i) {
      var x = c.colunas && c.colunas[col.id];
      if (x && x.ms > 0) { total += x.ms; partes.push({ col: col, i: i, ms: x.ms }); }
    });
    if (!total) return '<span class="muted">—</span>';
    var segs = partes.map(function (p) {
      return '<span class="seg" style="width:' + (p.ms / total * 100).toFixed(2) + '%;background:' + PALETA[p.i % PALETA.length] +
        '" title="' + esc(p.col.nome) + ': ' + R.fmtTempo(p.ms) + '"></span>';
    }).join('');
    var maior = partes.reduce(function (a, b) { return b.ms > a.ms ? b : a; });
    return '<div class="barra-col">' + segs + '</div><div class="mini">Maior: ' + esc(maior.col.nome) + ' (' + R.fmtTempo(maior.ms) + ')</div>';
  }

  // ---------- Desenho da tela ----------
  function deltaHtml(atual, ant, bomSobe) {
    if (ant == null) return '<span class="delta">sem comparação</span>';
    var d = atual - ant, seta = d > 0 ? '▲' : d < 0 ? '▼' : '●';
    var cls = (d === 0 || bomSobe === null) ? '' : ((d > 0) === bomSobe ? 'bom' : 'ruim');
    var iv = S.res && S.res.iv, dica = iv ? 'Período anterior: ' + R.fmtFull(iv.prevSince) + ' a ' + R.fmtFull(new Date(iv.since.getTime() - 1)) + ' (' + ant + ')' : '';
    return '<span class="delta ' + cls + '" title="' + esc(dica) + '">' + seta + ' ' + (d > 0 ? '+' : '') + d + ' <small>vs. período anterior (' + ant + ')</small></span>';
  }

  function renderKpis() {
    var k = S.res.kpis, p = k.prev;
    $('kpis').innerHTML =
      '<div class="kpi in clicavel" role="button" tabindex="0" data-tipo="in" title="Ver a lista de entradas"><div class="rot">' + ICONES.entrada + 'Entradas</div><div class="valor">' + k.in + '</div>' + deltaHtml(k.in, p && p.in, true) +
        '<div class="ver">Ver lista ›</div></div>' +
      '<div class="kpi out clicavel" role="button" tabindex="0" data-tipo="out" title="Ver a lista de saídas"><div class="rot">' + ICONES.saida + 'Saídas</div><div class="valor">' + k.out + '</div>' + deltaHtml(k.out, p && p.out, null) +
        '<div class="extra">Enviados a outro board</div><div class="ver">Ver lista ›</div></div>' +
      '<div class="kpi cancel clicavel" role="button" tabindex="0" data-tipo="cancel" title="Ver a lista de cancelamentos"><div class="rot">' + ICONES.cancelar + 'Cancelamentos</div><div class="valor">' + k.cancel + '</div>' + deltaHtml(k.cancel, p && p.cancel, false) +
        '<div class="extra">Taxa: <b>' + taxaTxt(k.taxa) + '</b> das entradas</div><div class="ver">Ver lista ›</div></div>' +
      '<div class="kpi saldo clicavel ' + (k.saldo < 0 ? 'neg' : k.saldo > 0 ? 'pos' : '') + '" role="button" tabindex="0" data-tipo="" title="Ver todas as movimentações"><div class="rot">' + ICONES.saldo + 'Saldo</div><div class="valor">' + (k.saldo > 0 ? '+' : '') + k.saldo + '</div>' +
        deltaHtml(k.saldo, p && p.saldo, true) + '<div class="extra">Entradas − saídas − cancelamentos</div><div class="ver">Ver lista ›</div></div>';
    var iv = S.res.iv;
    $('notaComparacao').textContent = p ? '"vs. período anterior" compara com os mesmos ' + Math.round(iv.span) + ' dias antes: ' + R.fmtFull(iv.prevSince) + ' a ' + R.fmtFull(new Date(iv.since.getTime() - 1)) + '.'
      : 'Sem comparação com período anterior neste período.';
  }

  function maxDe(lista, f) { return lista.reduce(function (m, x) { return Math.max(m, f(x)); }, 0); }

  function renderTudo() {
    var r = S.res;
    S.max = {
      in: maxDe(r.serie, function (x) { return x.in; }), out: maxDe(r.serie, function (x) { return x.out; }), cancel: maxDe(r.serie, function (x) { return x.cancel; }),
      destino: maxDe(r.destinos, function (x) { return x.qtd; }), etapa: maxDe(r.etapas, function (x) { return x.qtd; }),
      media: maxDe(r.gargalos, function (x) { return x.media; })
    };
    $('subtitulo').textContent = (S.boardNome ? S.boardNome + ' · ' : '') + r.iv.texto;
    renderKpis();
    $('tituloSerie').textContent = 'Evolução por ' + (r.serieRotulo === 'Dia' ? 'dia' : 'semana');
    var plural = function (n, um, varios) { return n + ' ' + (n === 1 ? um : varios); };
    $('resSerie').textContent = plural(r.serie.length, r.serieRotulo === 'Dia' ? 'dia' : 'semana', r.serieRotulo === 'Dia' ? 'dias' : 'semanas') + ' com movimentação';
    $('resEtiquetas').textContent = plural(r.etiquetas.length, 'etiqueta', 'etiquetas');
    $('resDestinos').textContent = plural(r.destinos.length, 'board', 'boards') + ' · ' + plural(r.kpis.out, 'cliente', 'clientes');
    $('resEtapas').textContent = plural(r.etapas.length, 'etapa', 'etapas') + ' · ' + plural(r.kpis.cancel, 'cancelamento', 'cancelamentos');
    S.tabelas.serie.set(r.serie);
    S.tabelas.etiquetas.set(r.etiquetas);
    S.tabelas.destinos.set(r.destinos);
    S.tabelas.etapas.set(r.etapas);
    S.tabelas.gargalos.set(r.gargalos);
    S.tabelas.clientes.set(r.clientes);
    S.tabelas.cards.set(r.linhas);
    aplicarFiltroTipo();
    S.tabelas.ignorados.set(r.ignorados);
    S.tabelas.pendentes.set(r.pendentes);
    S.tabelas.decididos.set(r.decididos);
    $('nPend').textContent = r.pendentes.length;
    $('nPend').parentNode.classList.toggle('alerta', r.pendentes.length > 0);
    $('contaDecisoes').textContent = 'Decisões salvas no Trello: ' + Object.keys(S.decisoes).length + ' (cabem cerca de 130).';
    $('naoContados').textContent = 'Fora das contagens neste período: ' + r.naoContados.arquivados + ' arquivado(s) e ' + r.naoContados.excluidos +
      ' excluído(s). ' + r.ignorados.length + ' evento(s) ignorado(s) por engano ou repetição (aba Ignorados).';
    $('nClientes').textContent = r.clientes.length;
    $('nCards').textContent = r.linhas.length;
    $('nIgn').textContent = r.ignorados.length;
    $('legendaColunas').innerHTML = r.colunas.length ? 'Tempo por coluna: ' + r.colunas.map(function (c, i) {
      return '<span><span class="ponto" style="background:' + PALETA[i % PALETA.length] + '"></span>' + esc(c.nome) + '</span>';
    }).join('') : '';
    var g = r.gargalos[0];
    $('destaqueGargalo').innerHTML = g && g.media > 0
      ? '<div class="destaque">Maior gargalo: <b>' + esc(g.nome) + '</b>, com tempo médio de ' + R.fmtTempo(g.media) + ' (' + g.n + ' cliente' + (g.n === 1 ? '' : 's') + ').</div>' : '';
    atualizarBadgeFiltros();
  }

  // ---------- Lista filtrada por tipo (clique nos quadros de números) ----------
  function aplicarFiltroTipo() {
    Array.prototype.forEach.call(document.querySelectorAll('#filtroTipo .chip'), function (b) {
      b.setAttribute('aria-pressed', b.getAttribute('data-t') === S.filtroTipo ? 'true' : 'false');
    });
    S.tabelas.cards.filtrar(S.filtroTipo ? function (r) { return r.tipo === S.filtroTipo; } : null);
  }
  function verLista(tipo) {
    S.filtroTipo = tipo || '';
    $('buscaCards').value = ''; S.tabelas.cards.busca('');
    aplicarFiltroTipo();
    mostrarAba('cards'); salvarPrefs();
  }

  // ---------- Abas, filtros e preferências ----------
  function mostrarAba(nome) {
    S.prefs.aba = nome;
    Array.prototype.forEach.call(document.querySelectorAll('.aba'), function (b) { b.setAttribute('aria-selected', b.getAttribute('data-aba') === nome ? 'true' : 'false'); });
    ['resumo', 'gargalos', 'clientes', 'cards', 'pendentes', 'ignorados'].forEach(function (a) { $('aba-' + a).hidden = a !== nome; });
  }
  function alternarFiltros(abrir) {
    var mostrar = abrir === undefined ? $('filtros').hidden : abrir;
    $('filtros').hidden = !mostrar;
    $('btnFiltros').setAttribute('aria-expanded', mostrar ? 'true' : 'false');
  }
  function atualizarBadgeFiltros() {
    var n = (S.prefs.etiqueta ? 1 : 0) + (S.cfg.janela ? 0 : 1) + (S.cfg.dedup ? 0 : 1);
    $('nFiltros').hidden = !n; $('nFiltros').textContent = n;
  }
  function marcarPeriodo() {
    Array.prototype.forEach.call(document.querySelectorAll('.chip'), function (b) { b.setAttribute('aria-pressed', b.getAttribute('data-p') === S.prefs.periodo ? 'true' : 'false'); });
    $('custom').hidden = S.prefs.periodo !== 'custom';
  }
  function salvarPrefs() {
    t.set('member', 'private', 'prefs', S.prefs).catch(function () {});
    t.set('board', 'shared', { listaDesist: S.cfg.listaDesist, janela: S.cfg.janela ? '5' : '0', dedup: S.cfg.dedup ? '1' : '0' }).catch(function () {});
  }

  // ---------- Decisões (ficam salvas no Trello e valem para a equipe toda) ----------
  var MAX_DECISOES = 3600; // limite do Trello é 4096 caracteres por tipo de dado
  function lerDecisoes(txt) {
    var m = {};
    String(txt || '').split(',').forEach(function (p) { var x = p.split(':'); if (x[0] && (x[1] === 'c' || x[1] === 'n')) m[x[0]] = x[1]; });
    return m;
  }
  function serializarDecisoes(m) {
    var limite = Date.now() / 1000 - 540 * 86400; // decisões com mais de ~18 meses saem
    return Object.keys(m).filter(function (id) { return parseInt(id.slice(0, 8), 16) > limite; })
      .map(function (id) { return id + ':' + m[id]; }).join(',');
  }
  async function decidir(id, dec) {
    var antes = S.decisoes[id];
    if (dec === 'x') delete S.decisoes[id]; else S.decisoes[id] = dec;
    var txt = serializarDecisoes(S.decisoes);
    if (txt.length > MAX_DECISOES) {
      if (antes) S.decisoes[id] = antes; else delete S.decisoes[id];
      toast('Limite de decisões salvas atingido. Desfaça algumas decisões antigas na aba Pendentes.');
      return;
    }
    try { await t.set('board', 'shared', 'decisoes', txt); }
    catch (e) {
      if (antes) S.decisoes[id] = antes; else delete S.decisoes[id];
      toast('Não consegui salvar a decisão no Trello: ' + e.message);
      return;
    }
    await recalcular();
  }

  // ---------- Exportar ----------
  function filtrosTxt() {
    var et = S.prefs.etiqueta === '' ? 'todas' : S.prefs.etiqueta === '__none__' ? 'sem etiqueta' :
      ((S.labels.filter(function (l) { return l.id === S.prefs.etiqueta; })[0] || {}).name || 'selecionada');
    var d = S.listas.filter(function (l) { return l.id === S.cfg.listaDesist; })[0];
    return 'Etiqueta: ' + et + ' · Desistentes: ' + (d ? d.name : 'não definida') +
      ' · Enganos: ' + (S.cfg.janela ? 'ignorar até 5 min' : 'não ignorar') + ' · Nº repetido: ' + (S.cfg.dedup ? 'conta 1x por mês' : 'não verificar');
  }
  function dadosExcel() {
    var r = S.res;
    return {
      titulo: 'Entradas x Saídas de clientes', boardNome: S.boardNome, periodoTxt: r.iv.texto, filtrosTxt: filtrosTxt(), geradoEm: new Date(),
      kpis: r.kpis, serieRotulo: r.serieRotulo,
      serie: r.serie.map(function (s) { return { rotulo: R.fmt(new Date(s.k)), in: s.in, out: s.out, cancel: s.cancel, saldo: s.saldo }; }),
      etiquetas: r.etiquetas, destinos: r.destinos, etapas: r.etapas, colunas: r.colunas, clientes: r.clientes.slice().sort(function (a, b) { return b.saida - a.saida; }), gargalos: r.gargalos,
      cards: r.linhas,
      pendentes: r.pendentes.map(function (ev) { return { date: ev.when, tipo: ev.tipo, card: ev.card.name || '(sem nome)', motivo: ev.c.motivo, porque: ev.pend.join(' · ') }; }),
      ignorados: r.ignorados.map(function (ev) { return { date: ev.when, tipo: ev.tipo, card: ev.card.name || '(sem nome)', motivo: ev.c.motivo, porque: ev.ign }; }),
      naoContados: r.naoContados
    };
  }
  async function exportar() {
    if (!S.res) return;
    var nome = 'fluxo-clientes-' + R.isoDia(new Date()) + '.xlsx', d = dadosExcel();
    try {
      await X.baixarBonito(d, nome);
      toast('Planilha gerada.');
    } catch (e) {
      try { X.baixarSimples(d, nome); toast('Não consegui gerar a planilha formatada (' + e.message + '). Baixei a versão simples.'); }
      catch (e2) { toast('Não consegui gerar o Excel: ' + e2.message); }
    }
  }

  // ---------- Início ----------
  function preencherSelects() {
    ['listaDesist', 'etiqueta'].forEach(function (id) { while ($(id).options.length > 1) $(id).remove(1); });
    S.listas.forEach(function (l) {
      var o = document.createElement('option'); o.value = l.id; o.textContent = l.name;
      $('listaDesist').appendChild(o);
    });
    S.labels.forEach(function (l) {
      var o = document.createElement('option'); o.value = l.id; o.textContent = l.name || (l.color ? '(' + l.color + ', sem nome)' : '(sem nome)');
      $('etiqueta').appendChild(o);
    });
    var sem = document.createElement('option'); sem.value = '__none__'; sem.textContent = 'Sem etiqueta'; $('etiqueta').appendChild(sem);
  }

  var iniciado = false;
  async function iniciar() {
    var api = t.getRestApi();
    var ok = await api.isAuthorized();
    $('auth').hidden = ok; $('app').hidden = !ok;
    if (!ok) return;
    S.token = await api.getToken();
    S.boardId = t.getContext().board;
    try { var b = await t.board('name'); S.boardNome = (b && b.name) || ''; } catch (e) {}
    S.listas = await t.lists('id', 'name');
    S.labels = await T.fetchLabels(S.boardId, S.token);

    var prefs = await t.get('member', 'private', 'prefs', null);
    if (prefs) Object.keys(S.prefs).forEach(function (k) { if (prefs[k] !== undefined) S.prefs[k] = prefs[k]; });
    S.cfg.listaDesist = await t.get('board', 'shared', 'listaDesist', '');
    S.cfg.janela = (await t.get('board', 'shared', 'janela', '5')) !== '0';
    S.cfg.dedup = (await t.get('board', 'shared', 'dedup', '1')) !== '0';
    S.decisoes = lerDecisoes(await t.get('board', 'shared', 'decisoes', ''));

    var existe = function (id) { return S.listas.some(function (l) { return l.id === id; }); };
    if (!S.cfg.listaDesist || !existe(S.cfg.listaDesist)) {
      var achada = S.listas.filter(function (l) { return /desist/i.test(l.name); })[0];
      S.cfg.listaDesist = achada ? achada.id : '';
    }
    if (S.prefs.etiqueta && S.prefs.etiqueta !== '__none__' && !S.labels.some(function (l) { return l.id === S.prefs.etiqueta; })) S.prefs.etiqueta = '';
    if (S.prefs.periodo === 'custom' && !S.prefs.de) S.prefs.periodo = '30';

    preencherSelects();
    $('etiqueta').value = S.prefs.etiqueta; $('listaDesist').value = S.cfg.listaDesist;
    $('janela').checked = S.cfg.janela; $('dedup').checked = S.cfg.dedup;
    $('de').value = S.prefs.de; $('ate').value = S.prefs.ate;
    $('btnAtualizar').innerHTML = ICONES.atualizar; $('btnFechar').innerHTML = ICONES.fechar;
    criarTabelas();
    ['serie', 'etiquetas', 'destinos', 'etapas'].forEach(function (k) {
      var det = $('det-' + k); det.open = !!(S.prefs.abertos && S.prefs.abertos[k]);
      det.addEventListener('toggle', function () { S.prefs.abertos = S.prefs.abertos || {}; S.prefs.abertos[k] = det.open; salvarPrefs(); });
    });
    marcarPeriodo(); mostrarAba(S.prefs.aba);
    atualizarBadgeFiltros();
    await carregar(false);
  }

  function ligarEventos() {
    $('btnAuth').onclick = function () { t.getRestApi().authorize({ scope: 'read', expiration: 'never' }).then(function () { iniciar(); }); };
    $('btnFechar').onclick = function () { t.closeModal(); };
    $('btnAtualizar').onclick = function () { carregar(true); };
    $('btnExportar').onclick = exportar;
    $('btnFiltros').onclick = function () { alternarFiltros(); };

    $('chips').addEventListener('click', function (e) {
      var b = e.target.closest('.chip'); if (!b) return;
      S.prefs.periodo = b.getAttribute('data-p');
      if (S.prefs.periodo === 'custom' && !S.prefs.de) {
        S.prefs.ate = R.isoDia(new Date()); S.prefs.de = R.isoDia(new Date(Date.now() - 30 * 86400000));
        $('de').value = S.prefs.de; $('ate').value = S.prefs.ate;
      }
      marcarPeriodo(); salvarPrefs(); carregar(false);
    });
    $('de').onchange = $('ate').onchange = function () {
      S.prefs.de = $('de').value; S.prefs.ate = $('ate').value; salvarPrefs();
      if (S.prefs.periodo === 'custom') carregar(false);
    };
    $('etiqueta').onchange = function () { S.prefs.etiqueta = $('etiqueta').value; salvarPrefs(); recalcular(); };
    $('listaDesist').onchange = function () { S.cfg.listaDesist = $('listaDesist').value; salvarPrefs(); recalcular(); };
    $('janela').onchange = function () { S.cfg.janela = $('janela').checked; salvarPrefs(); recalcular(); };
    $('dedup').onchange = function () { S.cfg.dedup = $('dedup').checked; salvarPrefs(); recalcular(); };
    $('kpis').addEventListener('click', function (e) {
      var k = e.target.closest('.kpi[data-tipo]'); if (k) verLista(k.getAttribute('data-tipo'));
    });
    $('kpis').addEventListener('keydown', function (e) {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      var k = e.target.closest('.kpi[data-tipo]'); if (k) { e.preventDefault(); verLista(k.getAttribute('data-tipo')); }
    });
    $('filtroTipo').addEventListener('click', function (e) {
      var b = e.target.closest('.chip'); if (!b) return;
      S.filtroTipo = b.getAttribute('data-t'); aplicarFiltroTipo();
    });
    $('abas').addEventListener('click', function (e) {
      var b = e.target.closest('.aba'); if (!b) return;
      mostrarAba(b.getAttribute('data-aba')); salvarPrefs();
    });
    $('abas').addEventListener('keydown', function (e) {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      var abas = Array.prototype.slice.call(document.querySelectorAll('.aba')), i = abas.indexOf(document.activeElement);
      if (i < 0) return;
      var prox = abas[(i + (e.key === 'ArrowRight' ? 1 : abas.length - 1)) % abas.length];
      prox.focus(); prox.click(); e.preventDefault();
    });
    document.addEventListener('click', function (e) {
      var b = e.target.closest('[data-dec]'); if (!b || b.disabled) return;
      b.disabled = true; decidir(b.getAttribute('data-id'), b.getAttribute('data-dec'));
    });
    $('buscaClientes').oninput = function () { S.tabelas.clientes && S.tabelas.clientes.busca(this.value); };
    $('buscaCards').oninput = function () { S.tabelas.cards && S.tabelas.cards.busca(this.value); };
  }

  ligarEventos();
  t.render(function () {
    aplicarTema();
    if (iniciado) return;
    iniciado = true;
    return iniciar();
  });
})();
