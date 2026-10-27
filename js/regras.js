/* Regras de contagem: entradas, saídas, cancelamentos, enganos, tempo por coluna e análises.
   Não mexe na tela nem no Trello: só recebe dados e devolve resultados (por isso dá para testar). */
(function (root) {
  'use strict';
  var R = {};

  // ---------- Formatação ----------
  var pad = function (n) { return String(n).padStart(2, '0'); };
  R.pad = pad;
  R.fmt = function (d) { return pad(d.getDate()) + '/' + pad(d.getMonth() + 1); };
  R.fmtFull = function (d) { return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear(); };
  R.fmtHora = function (d) { return pad(d.getHours()) + ':' + pad(d.getMinutes()); };
  R.isoDia = function (d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); };
  R.esc = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  };

  // Tempo legível: minutos, horas ou dias
  R.fmtTempo = function (ms) {
    var min = Math.round(ms / 60000);
    if (min < 1) return ms > 0 ? '< 1 min' : '0 min';
    if (min < 60) return min + ' min';
    var h = Math.floor(min / 60), m = min % 60;
    if (h < 24) return h + ' h' + (m ? ' ' + m + ' min' : '');
    var d = Math.floor(h / 24); h = h % 24;
    return d + (d === 1 ? ' dia' : ' dias') + (h ? ' ' + h + ' h' : '');
  };
  R.horas = function (ms) { return Math.round(ms / 360000) / 10; };

  // "1234567890 - NOME" -> { numero: "1234567890", nome: "NOME" }
  R.partirNome = function (n) {
    var m = String(n || '').match(/^\s*(\d+)\s*(?:[-–—:]\s*)?(.*)$/);
    if (m) return { numero: m[1], nome: m[2].trim() };
    return { numero: '', nome: String(n || '').trim() };
  };

  // A data de criação do card está escondida nos 8 primeiros caracteres do ID
  R.criacaoCard = function (id) { return new Date(parseInt(String(id).substring(0, 8), 16) * 1000); };

  R.nomeBoard = function (b, nomes) {
    if (!b) return 'undefined';
    return b.name || (nomes && nomes[b.id]) || 'undefined';
  };

  // ---------- Classificação de cada ação do Trello ----------
  // tipo: in (entrada) | out (saiu para outro board) | cancel (foi para desistentes)
  //       rem (arquivado/excluído: não conta) | unarq (desarquivado) | saiu_desist (saiu de desistentes)
  R.classify = function (a, cfg) {
    var d = a.data || {};
    var lista = (d.list && d.list.name) || '';
    var nomes = cfg && cfg.boardNomes;
    switch (a.type) {
      case 'createCard': return { tipo: 'in', motivo: 'Criado', origem: '', destino: lista };
      case 'copyCard': return { tipo: 'in', motivo: 'Copiado', copia: true, origem: '', destino: lista };
      case 'moveCardToBoard':
        return { tipo: 'in', motivo: 'Veio de outro board: ' + R.nomeBoard(d.boardSource, nomes), origem: '', destino: lista };
      case 'moveCardFromBoard':
        var nb = R.nomeBoard(d.boardTarget, nomes);
        return { tipo: 'out', motivo: 'Enviado a outro board: ' + nb, origem: lista, destino: '', destinoBoard: nb };
      case 'deleteCard': return { tipo: 'rem', motivo: 'Excluído', sub: 'excluido', origem: lista, destino: '' };
      case 'updateCard':
        if (d.old && d.old.closed === false && d.card && d.card.closed === true)
          return { tipo: 'rem', motivo: 'Arquivado', sub: 'arquivado', origem: lista, destino: '' };
        if (d.old && d.old.closed === true && d.card && d.card.closed === false)
          return { tipo: 'unarq', motivo: 'Desarquivado', origem: '', destino: lista };
        if (cfg && cfg.listaDesist && d.listAfter && d.listAfter.id === cfg.listaDesist)
          return { tipo: 'cancel', motivo: 'Cancelamento: movido para ' + d.listAfter.name,
                   origem: (d.listBefore && d.listBefore.name) || '', destino: d.listAfter.name };
        if (cfg && cfg.listaDesist && d.listBefore && d.listBefore.id === cfg.listaDesist &&
            d.listAfter && d.listAfter.id !== cfg.listaDesist)
          return { tipo: 'saiu_desist', motivo: 'Saiu de desistentes', origem: d.listBefore.name, destino: d.listAfter.name };
        return null;
    }
    return null;
  };

  R.passaEtiqueta = function (cardId, etiqueta, cardLabels) {
    if (!etiqueta) return true;
    var ls = cardId ? cardLabels[cardId] : null;
    if (etiqueta === '__none__') return !ls || ls.length === 0;
    return !!ls && ls.indexOf(etiqueta) >= 0;
  };

  // Transforma as ações do Trello em eventos em ordem cronológica
  R.montarEventos = function (actions, cfg, etiqueta, cardLabels) {
    var eventos = [];
    actions.forEach(function (a) {
      var c = R.classify(a, cfg);
      if (!c) return;
      var card = (a.data && a.data.card) || {};
      if (!R.passaEtiqueta(card.id, etiqueta, cardLabels || {})) return;
      eventos.push({
        a: a, c: c, card: card, tipo: c.tipo, when: new Date(a.date),
        id: card.id || ('sem-id-' + a.id),
        criacao: card.id ? R.criacaoCard(card.id) : null,
        ign: null
      });
    });
    eventos.sort(function (x, y) { return x.when - y.when || (x.a.id < y.a.id ? -1 : 1); });
    return eventos;
  };

  // ---------- Duplicatas e enganos ----------
  var H24 = 24 * 3600000;
  var MOTIVO_PAR = 'Criado, copiado ou recebido e removido em poucos minutos (provável engano)';
  var MOTIVO_REM = 'Criado ou recebido e arquivado/excluído em poucos minutos (provável engano)';
  var MOTIVO_COPIA = 'Copiado e arquivado/excluído em até 24 h (provável cópia por engano)';
  var MOTIVO_VOLTOU = 'Voltou de outro board em até 24 h: não conta como nova entrada';
  var MOTIVO_VOLTOU_SAIDA = 'Voltou de outro board em até 24 h: a saída foi desfeita';
  var PEND_RETORNO = 'Voltou de outro board depois de mais de 24 h: é um projeto novo ou o mesmo projeto?';
  var MOTIVO_COPIA_INTERNA = 'Cópia feita dentro do próprio board: não conta como entrada';
  var PEND_COPIA_ORIGEM_SAIU = 'Cópia de um card que estava neste board e depois saiu ou foi excluído: foi cópia feita aqui ou veio de outro board?';
  var PEND_COPIA_INTERNA = 'Cópia feita no próprio board que depois saiu: é um projeto adicional?';
  var PEND_DOIS = 'Parece ser de 2 projetos (etiqueta 2pack, nome ou descrição): confirme se conta';
  var PEND_REPETIDO = 'Mesmo nº de cliente já contado no mês, mas o card parece ser de outro projeto: confirme se conta';

  function addPend(ev, txt) {
    if (!ev.pend) ev.pend = [];
    if (ev.pend.indexOf(txt) < 0) ev.pend.push(txt);
  }
  R.addPend = addPend;

  // Recebe os eventos em ordem cronológica e devolve só entradas, saídas e cancelamentos,
  // com "ign" (ignorado) e "pend" (pendente de confirmação) preenchidos quando for o caso.
  R.processar = function (lista, janelaMs) {
    var status = {}, ultEntrada = {}, ultSaida = {}, ultRem = {}, out = [];
    function dentro(ev, ref, lim) { return lim > 0 && ref && (ev.when - ref) <= lim; }

    lista.forEach(function (ev) {
      var id = ev.id, st = status[id];
      switch (ev.tipo) {
        case 'in':
          if (st === 'in') { // duplicado (ex.: createCard + copyCard do mesmo card)
            var ja = ultEntrada[id];
            if (ev.c && ev.c.copia && ja && ja.c && !ja.c.copia) { ja.c.copia = true; ja.c.motivo = 'Copiado'; }
            return;
          }
          var voltou = st === 'out' && ultSaida[id] && String(ev.c.motivo || '').indexOf('Veio de outro board') === 0;
          if (voltou) {
            var sx = ultSaida[id];
            if (janelaMs > 0 && (ev.when - sx.when) <= H24) {
              ev.ign = MOTIVO_VOLTOU; sx.ign = MOTIVO_VOLTOU_SAIDA;
              status[id] = 'in'; out.push(ev); // mantém a entrada original como referência de tempo
              break;
            }
            addPend(ev, PEND_RETORNO);
          }
          if (ev.c && ev.c.copia && ev.c.copiaInterna) ev.ign = MOTIVO_COPIA_INTERNA;
          else if (ev.c && ev.c.copia && ev.c.copiaOrigemSaiu) addPend(ev, PEND_COPIA_ORIGEM_SAIU);
          status[id] = 'in'; ultEntrada[id] = ev; out.push(ev);
          break;

        case 'out':
        case 'cancel':
          if (st === 'out' || st === 'cancel') return; // já tinha saído ou cancelado
          status[id] = ev.tipo;
          var ent = ultEntrada[id];
          ev.ref = ent ? ent.when : (ev.criacao || null);
          var copia = ent && ent.c && ent.c.copia;
          if (ent && ent.c && ent.c.copiaInterna && ent.ign === MOTIVO_COPIA_INTERNA) { ent.ign = null; addPend(ent, PEND_COPIA_INTERNA); }
          // cópia enviada a outro board é um projeto adicional legítimo: não entra na regra de engano
          if (!(copia && ev.tipo === 'out') && dentro(ev, ev.ref, janelaMs)) {
            ev.ign = MOTIVO_PAR;
            if (ent && !ent.ign) { ent.ign = MOTIVO_PAR; ev.pairEntry = ent; }
          }
          ultSaida[id] = ev; out.push(ev);
          break;

        case 'rem': // arquivado/excluído não conta como saída; só serve para detectar engano
          if (st === 'rem' || st === 'out' || st === 'cancel') return;
          status[id] = 'rem';
          var e2 = ultEntrada[id];
          var ehCopia = e2 && e2.c && e2.c.copia;
          var lim = ehCopia && janelaMs > 0 ? H24 : janelaMs;
          if (e2 && !e2.ign && dentro(ev, e2.when, lim)) {
            e2.ign = ehCopia ? MOTIVO_COPIA : MOTIVO_REM;
            ultRem[id] = { pairEntry: e2 };
          }
          break;

        case 'unarq':
          if (st !== 'rem') return;
          status[id] = 'in';
          if (ultRem[id] && ultRem[id].pairEntry) { ultRem[id].pairEntry.ign = null; ultRem[id] = null; } // o card existe
          break;

        case 'saiu_desist':
          return; // o cancelamento continua valendo e não gera nova entrada
      }
    });
    return out;
  };

  // ---------- Decisões tomadas por você ----------
  // decisoes: { idDaAcao: 'c' (conta) | 'n' (não conta) }
  R.aplicarDecisoes = function (kept, decisoes) {
    kept.forEach(function (ev) {
      var d = decisoes && ev.a && decisoes[ev.a.id];
      if (!d) return;
      ev.decidido = d;
      ev.pend = null;
      ev.ign = d === 'c' ? null : 'Você marcou como "não contar"';
    });
    return kept;
  };

  // Lê se o card parece ser de 2 projetos.
  // Nome: "projeto_1"/"projeto_2", "2 projetos", "segundo projeto", "2pack". Descrição: só frases claras
  // (para não marcar qualquer "projeto 1" escrito em um briefing comum). Etiqueta: 2pack.
  R.leDoisProjetos = function (nomes, etiquetas, descs) {
    var et = (etiquetas || []).some(function (n) { return /(^|[^a-z0-9])2\s*[-_ ]?pack([^a-z]|$)/i.test(n || ''); });
    if (et) return true;
    var nome = (nomes || []).filter(Boolean).join(' \n ');
    var desc = (descs || []).filter(Boolean).join(' \n ');
    if (/projeto[\s_\-]*0?[12]\b/i.test(nome)) return true;
    var tudo = nome + ' \n ' + desc;
    return /\b(2|dois)\s+projetos\b/i.test(tudo) || /\bsegundo\s+projeto\b/i.test(tudo) ||
      /(^|[^a-z0-9])2\s*[-_ ]?pack([^a-z]|$)/i.test(tudo) || /projeto[\s_\-]*0?2\b/i.test(desc);
  };

  // Marca o que precisa da sua confirmação (já fica contado provisoriamente)
  R.marcarPendencias = function (kept) {
    kept.forEach(function (ev) {
      if (ev.ign || ev.decidido || ev.tipo !== 'in') return;
      if (ev.doisProjetos && !(ev.pend && ev.pend.indexOf(PEND_REPETIDO) >= 0)) addPend(ev, PEND_DOIS);
    });
    return kept;
  };

  // ---------- Mesmo número de cliente ----------
  // O número do cliente é o que vem no início do nome do card ("1234567890 - NOME").
  // O mesmo número só é contado uma vez por tipo (entrada, saída, cancelamento) em cada mês,
  // mesmo que apareça em cards diferentes. Quem já foi ignorado por outro motivo não ocupa a vaga.
  R.NUM_MIN = 6;
  R.deduplicarNumeros = function (kept) {
    var vistos = {};
    kept.forEach(function (ev) {
      if (ev.decidido === 'n') return;
      if (ev.ign && ev.decidido !== 'c') return;
      if (ev.pend && ev.pend.indexOf(PEND_RETORNO) >= 0) return; // retorno de outro board: você decide, não é repetição
      var numero = R.partirNome(ev.card && ev.card.name).numero;
      if (!numero || numero.length < R.NUM_MIN) return; // sem número (ou número curto demais): não dá para comparar com segurança
      var mes = ev.when.getFullYear() + '-' + R.pad(ev.when.getMonth() + 1);
      var chave = ev.tipo + '|' + numero + '|' + mes;
      if (vistos[chave] && ev.decidido !== 'c') {
        if (ev.doisProjetos) {
          addPend(ev, PEND_REPETIDO); // parece outro projeto do mesmo cliente: você confirma
        } else {
          ev.ign = 'Nº de cliente ' + numero + ' repetido no mês: já foi contado uma vez';
          ev.repetido = true;
        }
      } else {
        vistos[chave] = true;
      }
    });
    return kept;
  };

  // ---------- Tempo em cada coluna ----------
  // acts: ações do card em ordem cronológica { type, date: Date, data }. ref = entrada, fim = saída.
  R.tempoPorColuna = function (acts, ref, fim, listaFinal) {
    var tempos = {}, atual = null, desde = null, pausado = false;
    function fecha(ate) {
      if (atual && desde && !pausado && ate > desde) {
        tempos[atual.id] = tempos[atual.id] || { nome: atual.nome, ms: 0 };
        tempos[atual.id].ms += ate - desde;
      }
    }
    acts.forEach(function (a) {
      var d = a.data || {}, quando = a.date;
      if (quando > fim) return;
      if (a.type === 'createCard' || a.type === 'copyCard' || a.type === 'moveCardToBoard') {
        fecha(quando);
        atual = d.list ? { id: d.list.id, nome: d.list.name } : null;
        desde = quando; pausado = false;
      } else if (a.type === 'moveCardFromBoard') {
        fecha(quando); atual = null; desde = null;
      } else if (a.type === 'updateCard') {
        if (d.listAfter) {
          if (!atual && !desde && d.listBefore) { atual = { id: d.listBefore.id, nome: d.listBefore.name }; desde = ref; }
          fecha(quando);
          atual = { id: d.listAfter.id, nome: d.listAfter.name }; desde = quando;
        } else if (d.old && d.old.closed === false && d.card && d.card.closed === true) {
          fecha(quando); pausado = true; // arquivado: o tempo para de contar
        } else if (d.old && d.old.closed === true && d.card && d.card.closed === false) {
          pausado = false; desde = quando; // desarquivado: volta a contar
        }
      }
    });
    fecha(fim);
    if (!Object.keys(tempos).length && listaFinal && listaFinal.id && ref && fim > ref) {
      tempos[listaFinal.id] = { nome: listaFinal.name, ms: fim - ref }; // sem histórico: usa a lista em que estava
    }
    return tempos;
  };

  // ---------- Clientes (quem saiu ou cancelou) ----------
  R.montarClientes = function (contados, cardLabels, labelsPorId) {
    var lista = [];
    contados.forEach(function (ev) {
      if (ev.tipo !== 'out' && ev.tipo !== 'cancel') return;
      var ref = ev.ref || ev.criacao;
      if (!ref) return;
      var p = R.partirNome(ev.card.name);
      var ids = (cardLabels && cardLabels[ev.id]) || [];
      lista.push({
        id: ev.id,
        situacao: ev.tipo,
        numero: p.numero, nome: p.nome, card: ev.card.name || '(sem nome)',
        entrada: ref, saida: ev.when, ms: Math.max(0, ev.when - ref),
        destino: ev.tipo === 'out' ? ev.c.destinoBoard : ev.c.destino,
        etapa: ev.c.origem || '',
        motivo: ev.c.motivo,
        etiquetas: ids,
        shortLink: ev.card.shortLink || '',
        listaFinal: (ev.a && ev.a.data && ev.a.data.list) || null,
        colunas: {}
      });
    });
    return lista;
  };

  // Aplica o histórico de movimentações (por card) aos clientes
  R.aplicarColunas = function (clientes, porCard) {
    clientes.forEach(function (c) {
      var acts = (porCard[c.id] || []).slice().sort(function (x, y) { return x.date - y.date || (x.id < y.id ? -1 : 1); });
      c.colunas = R.tempoPorColuna(acts, c.entrada, c.saida, c.listaFinal);
    });
  };

  // Colunas (listas) que aparecem, na ordem em que estão no board
  R.colunasUsadas = function (clientes, listasBoard) {
    var usados = {};
    clientes.forEach(function (c) {
      Object.keys(c.colunas || {}).forEach(function (id) { if (c.colunas[id].ms > 0) usados[id] = c.colunas[id].nome; });
    });
    var ordem = [];
    (listasBoard || []).forEach(function (l) { if (usados[l.id]) { ordem.push({ id: l.id, nome: usados[l.id] }); delete usados[l.id]; } });
    Object.keys(usados).forEach(function (id) { ordem.push({ id: id, nome: usados[id] }); });
    return ordem;
  };

  // ---------- Análises ----------
  R.contar = function (evs) {
    var n = { in: 0, out: 0, cancel: 0 };
    evs.forEach(function (ev) { if (!ev.ign && n[ev.tipo] !== undefined) n[ev.tipo]++; });
    n.saldo = n.in - n.out - n.cancel;
    n.taxa = n.in > 0 ? n.cancel / n.in : null;
    return n;
  };

  R.bucketKey = function (date, span) {
    var d = new Date(date);
    if (span > 31) { var diff = (d.getDay() + 6) % 7; d.setDate(d.getDate() - diff); } // semana (segunda-feira)
    d.setHours(0, 0, 0, 0);
    return d.getTime();
  };

  R.serie = function (counted, span) {
    var b = {};
    counted.forEach(function (ev) {
      var k = R.bucketKey(ev.when, span);
      b[k] = b[k] || { k: k, in: 0, out: 0, cancel: 0 };
      b[k][ev.tipo]++;
    });
    return Object.keys(b).map(Number).sort(function (x, y) { return y - x; }).map(function (k) {
      var o = b[k]; o.saldo = o.in - o.out - o.cancel; return o;
    });
  };

  R.destinos = function (counted) {
    var m = {};
    counted.forEach(function (ev) { if (ev.tipo === 'out') { var n = ev.c.destinoBoard || 'undefined'; m[n] = (m[n] || 0) + 1; } });
    return Object.keys(m).map(function (n) { return { nome: n, qtd: m[n] }; }).sort(function (a, b) { return b.qtd - a.qtd; });
  };

  R.etapasDesistencia = function (counted) {
    var m = {};
    counted.forEach(function (ev) { if (ev.tipo === 'cancel') { var n = ev.c.origem || '(sem informação)'; m[n] = (m[n] || 0) + 1; } });
    return Object.keys(m).map(function (n) { return { nome: n, qtd: m[n] }; }).sort(function (a, b) { return b.qtd - a.qtd; });
  };

  // Entradas, saídas, cancelamentos e tempo médio por etiqueta (etiquetas que o card tem hoje)
  R.porEtiqueta = function (counted, clientes, cardLabels, labels) {
    var info = {};
    (labels || []).forEach(function (l) { info[l.id] = l; });
    var linhas = {};
    function linha(id) {
      if (!linhas[id]) {
        var l = info[id];
        linhas[id] = { id: id, nome: id === '__none__' ? 'Sem etiqueta' : (l ? (l.name || '(sem nome)') : '(etiqueta removida)'),
                       cor: l ? l.color : null, in: 0, out: 0, cancel: 0, msTotal: 0, nClientes: 0 };
      }
      return linhas[id];
    }
    function idsDe(cardId) {
      var ls = (cardLabels && cardLabels[cardId]) || [];
      return ls.length ? ls : ['__none__'];
    }
    counted.forEach(function (ev) {
      if (ev.ign || ev.tipo === 'rem') return;
      idsDe(ev.id).forEach(function (lid) { linha(lid)[ev.tipo]++; });
    });
    clientes.forEach(function (c) {
      idsDe(c.id).forEach(function (lid) { var l = linha(lid); l.msTotal += c.ms; l.nClientes++; });
    });
    return Object.keys(linhas).map(function (k) {
      var l = linhas[k];
      l.saldo = l.in - l.out - l.cancel;
      l.taxa = l.in > 0 ? l.cancel / l.in : null;
      l.tempoMedio = l.nClientes ? l.msTotal / l.nClientes : null;
      return l;
    }).sort(function (a, b) { return (b.in + b.out + b.cancel) - (a.in + a.out + a.cancel); });
  };

  // Em qual coluna os clientes ficam mais tempo
  R.gargalos = function (clientes, colunas) {
    var totalGeral = 0;
    var linhas = colunas.map(function (col) {
      var tempos = [];
      clientes.forEach(function (c) { var x = c.colunas && c.colunas[col.id]; if (x && x.ms > 0) tempos.push(x.ms); });
      var total = tempos.reduce(function (s, v) { return s + v; }, 0);
      totalGeral += total;
      tempos.sort(function (a, b) { return a - b; });
      return {
        id: col.id, nome: col.nome, n: tempos.length, total: total,
        media: tempos.length ? total / tempos.length : 0,
        mediana: tempos.length ? tempos[Math.floor((tempos.length - 1) / 2)] : 0,
        maior: tempos.length ? tempos[tempos.length - 1] : 0
      };
    });
    linhas.forEach(function (l) { l.share = totalGeral ? l.total / totalGeral : 0; });
    return linhas.sort(function (a, b) { return b.media - a.media; });
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = R; else root.Regras = R;
})(typeof window !== 'undefined' ? window : this);
