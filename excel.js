/* Exportação para Excel. Versão bonita (ExcelJS) e versão simples (SheetJS) como reserva. */
(function (root) {
  'use strict';
  var X = {};
  var R = root.Regras || (typeof require !== 'undefined' ? require('./regras.js') : null);

  var COR = {
    navy: 'FF172B4D', azul: 'FF0C66E4', verde: 'FF1F845A', vermelho: 'FFC9372C', laranja: 'FFE56910', cinza: 'FF626F86',
    zebra: 'FFF7F8F9', borda: 'FFDFE1E6', branco: 'FFFFFFFF',
    verdeL: 'FFDCFFF1', azulL: 'FFE9F2FF', vermelhoL: 'FFFFECEB', cinzaL: 'FFF1F2F4'
  };
  var FMT_DATAHORA = 'dd/mm/yyyy hh:mm';
  var FMT_SALDO = '+0;[Red]-0;0';
  var SITUACAO = { out: 'Saiu para outro board', cancel: 'Cancelou (desistente)' };
  var TIPO = { in: 'Entrada', out: 'Saída', cancel: 'Cancelamento' };

  // O Excel guarda datas sem fuso; ajustamos para aparecer a hora local
  function loc(d) { return d ? new Date(d.getTime() - d.getTimezoneOffset() * 60000) : null; }
  function preencher(cell, argb) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }
  function bordas() {
    var l = { style: 'thin', color: { argb: COR.borda } };
    return { top: l, left: l, bottom: l, right: l };
  }
  function pctTxt(v) { return v == null ? '—' : (Math.round(v * 1000) / 10).toString().replace('.', ',') + '%'; }

  function titulo(ws, texto, sub, ncols) {
    ws.mergeCells(1, 1, 1, Math.max(ncols, 2));
    var c = ws.getCell(1, 1);
    c.value = texto;
    c.font = { name: 'Calibri', size: 16, bold: true, color: { argb: COR.navy } };
    c.alignment = { vertical: 'middle' };
    ws.getRow(1).height = 28;
    ws.mergeCells(2, 1, 2, Math.max(ncols, 2));
    var s = ws.getCell(2, 1);
    s.value = sub;
    s.font = { name: 'Calibri', size: 10, color: { argb: COR.cinza } };
    s.alignment = { vertical: 'middle', wrapText: true };
    ws.getRow(2).height = 30;
  }

  function cabecalho(ws, linha, nomes, colIni) {
    nomes.forEach(function (n, i) {
      var c = ws.getCell(linha, (colIni || 1) + i);
      c.value = n;
      c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COR.branco } };
      preencher(c, COR.navy);
      c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      c.border = bordas();
    });
    ws.getRow(linha).height = 32;
  }

  // Escreve as linhas com zebra, bordas e formatos. fmts[i] = formato numérico da coluna i.
  function linhasTabela(ws, linhaIni, linhas, fmts, aligns, colIni) {
    linhas.forEach(function (vals, r) {
      vals.forEach(function (v, i) {
        var c = ws.getCell(linhaIni + r, (colIni || 1) + i);
        c.value = v === undefined ? null : v;
        c.font = { name: 'Calibri', size: 10, color: { argb: 'FF172B4D' } };
        c.border = bordas();
        c.alignment = { vertical: 'middle', horizontal: (aligns && aligns[i]) || (typeof v === 'number' || v instanceof Date ? 'right' : 'left') };
        if (fmts && fmts[i]) c.numFmt = fmts[i];
        if (r % 2 === 1) preencher(c, COR.zebra);
      });
    });
  }

  function larguras(ws, cabecalhos, linhas, minimo, maximo, colIni) {
    cabecalhos.forEach(function (h, i) {
      var max = String(h).length * 0.9;
      linhas.forEach(function (l) {
        var v = l[i];
        var len = v instanceof Date ? 17 : String(v == null ? '' : v).length;
        if (len > max) max = len;
      });
      ws.getColumn((colIni || 1) + i).width = Math.min(maximo || 46, Math.max(minimo || 11, Math.ceil(max * 1.1) + 2));
    });
  }

  function paginaCfg(ws) {
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  }

  // Aba de dados: título, cabeçalho na linha 4, filtro automático, linha congelada
  function abaDados(wb, nome, cor, tit, sub, cab, linhas, o) {
    o = o || {};
    var ws = wb.addWorksheet(nome, {
      properties: { tabColor: { argb: cor } },
      views: [{ state: 'frozen', ySplit: 4, xSplit: o.congelarCols || 0, showGridLines: false }]
    });
    titulo(ws, tit, sub, Math.max(cab.length, 6));
    cabecalho(ws, 4, cab);
    if (!linhas.length) {
      ws.mergeCells(5, 1, 5, Math.max(cab.length, 2));
      var v = ws.getCell(5, 1);
      v.value = 'Nenhum registro neste período.';
      v.font = { name: 'Calibri', size: 10, italic: true, color: { argb: COR.cinza } };
      v.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getRow(5).height = 28;
    } else {
      linhasTabela(ws, 5, linhas, o.fmts, o.aligns);
      ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + linhas.length, column: cab.length } };
    }
    larguras(ws, cab, linhas, o.min, o.max);
    paginaCfg(ws);
    return ws;
  }

  function escala(ws, ref, corFim) {
    ws.addConditionalFormatting({
      ref: ref,
      rules: [{ type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'max' }], color: [{ argb: COR.branco }, { argb: corFim }] }]
    });
  }

  function colLetra(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  function deltaTxt(atual, ant, rotuloAnt) {
    if (ant == null) return 'sem comparação';
    var d = atual - ant;
    return (d > 0 ? '▲ +' : d < 0 ? '▼ ' : '● ') + d + ' vs. ' + rotuloAnt + ' (' + ant + ')';
  }

  // ---------- Versão bonita ----------
  X.construir = function (ExcelJS, d) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Fluxo de Clientes';
    wb.created = new Date();
    var sub = 'Período: ' + d.periodoTxt + '   |   Filtros: ' + d.filtrosTxt + '   |   Gerado em ' + R.fmtFull(d.geradoEm) + ' às ' + R.fmtHora(d.geradoEm);
    var k = d.kpis, p = k.prev;

    // ===== Resumo =====
    var ws = wb.addWorksheet('Resumo', { properties: { tabColor: { argb: COR.navy } }, views: [{ showGridLines: false }] });
    titulo(ws, d.titulo + (d.boardNome ? ' — ' + d.boardNome : ''), sub, 5);
    ['A', 'B', 'C', 'D', 'E'].forEach(function (l) { ws.getColumn(l).width = 24; });

    var cards = [
      { rot: 'ENTRADAS', val: k.in, cor: COR.verde, fundo: COR.verdeL, dl: deltaTxt(k.in, p && p.in, 'período anterior'), fmt: '0' },
      { rot: 'SAÍDAS (outro board)', val: k.out, cor: COR.azul, fundo: COR.azulL, dl: deltaTxt(k.out, p && p.out, 'período anterior'), fmt: '0' },
      { rot: 'CANCELAMENTOS', val: k.cancel, cor: COR.vermelho, fundo: COR.vermelhoL, dl: deltaTxt(k.cancel, p && p.cancel, 'período anterior'), fmt: '0' },
      { rot: 'SALDO (ativos)', val: k.saldo, cor: k.saldo < 0 ? COR.vermelho : COR.navy, fundo: COR.cinzaL, dl: deltaTxt(k.saldo, p && p.saldo, 'período anterior'), fmt: '+0;-0;0' },
      { rot: 'TAXA DE CANCELAMENTO', val: k.taxa, cor: COR.laranja, fundo: 'FFFFF4E5', dl: 'cancelamentos ÷ entradas', fmt: '0.0%' }
    ];
    cards.forEach(function (c, i) {
      var col = i + 1;
      var a = ws.getCell(4, col), b = ws.getCell(5, col), e = ws.getCell(6, col);
      a.value = c.rot; a.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COR.cinza } };
      b.value = c.val == null ? '—' : c.val; b.numFmt = c.fmt; b.font = { name: 'Calibri', size: 24, bold: true, color: { argb: c.cor } };
      e.value = c.dl; e.font = { name: 'Calibri', size: 9, color: { argb: COR.cinza } };
      [a, b, e].forEach(function (x) {
        preencher(x, c.fundo);
        x.alignment = { horizontal: 'center', vertical: 'middle' };
        x.border = { left: { style: 'medium', color: { argb: COR.branco } }, right: { style: 'medium', color: { argb: COR.branco } } };
      });
    });
    ws.getRow(4).height = 20; ws.getRow(5).height = 40; ws.getRow(6).height = 20;

    var t0 = 8;
    ws.getCell(t0, 1).value = 'Evolução por ' + (d.serieRotulo === 'Semana de' ? 'semana' : 'dia');
    ws.getCell(t0, 1).font = { name: 'Calibri', size: 12, bold: true, color: { argb: COR.navy } };
    cabecalho(ws, t0 + 1, [d.serieRotulo, 'Entradas', 'Saídas', 'Cancelamentos', 'Saldo']);
    var serie = d.serie.map(function (s) { return [s.rotulo, s.in, s.out, s.cancel, s.saldo]; });
    serie.push(['TOTAL', k.in, k.out, k.cancel, k.saldo]);
    linhasTabela(ws, t0 + 2, serie, [null, '0', '0', '0', FMT_SALDO], ['left', 'right', 'right', 'right', 'right']);
    var tot = t0 + 2 + serie.length - 1;
    for (var c = 1; c <= 5; c++) {
      var cel = ws.getCell(tot, c);
      cel.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COR.navy } };
      preencher(cel, COR.azulL);
    }
    var nota = tot + 2;
    ws.mergeCells(nota, 1, nota, 5);
    ws.getCell(nota, 1).value = 'Não entram nas contagens: ' + d.naoContados.arquivados + ' arquivado(s) e ' + d.naoContados.excluidos +
      ' excluído(s) no período; ' + d.ignorados.length + ' evento(s) ignorado(s) por engano ou repetição (aba "Ignorados"); ' + d.pendentes.length + ' pendente(s) de confirmação, contados provisoriamente (aba "Pendentes").';
    ws.getCell(nota, 1).font = { name: 'Calibri', size: 9, italic: true, color: { argb: COR.cinza } };
    ws.getCell(nota, 1).alignment = { wrapText: true, vertical: 'top' };
    ws.getRow(nota).height = 28;
    paginaCfg(ws);

    // ===== Por etiqueta =====
    var linhasEt = d.etiquetas.map(function (e) {
      return [e.nome, e.in, e.out, e.cancel, e.taxa, e.saldo, e.nClientes, e.tempoMedio == null ? '' : R.fmtTempo(e.tempoMedio), e.tempoMedio == null ? '' : R.horas(e.tempoMedio)];
    });
    abaDados(wb, 'Por etiqueta', COR.verde, 'Resultado por etiqueta', sub + '   |   Etiquetas que o card tem hoje; cards excluídos aparecem em "Sem etiqueta".',
      ['Etiqueta', 'Entradas', 'Saídas', 'Cancelamentos', 'Taxa de cancelamento', 'Saldo', 'Clientes que saíram ou cancelaram', 'Tempo médio', 'Tempo médio (h)'],
      linhasEt, { fmts: [null, '0', '0', '0', '0.0%', FMT_SALDO, '0', null, '0.0'], aligns: ['left', 'right', 'right', 'right', 'right', 'right', 'right', 'right', 'right'], min: 12 });

    // ===== Gargalos =====
    var linhasG = d.gargalos.map(function (g) {
      return [g.nome, g.n, R.fmtTempo(g.media), R.horas(g.media), R.fmtTempo(g.mediana), R.fmtTempo(g.maior), R.horas(g.maior), R.horas(g.total), g.share];
    });
    var wsG = abaDados(wb, 'Gargalos', COR.laranja, 'Onde os clientes ficam mais tempo', sub + '   |   Considera os clientes que saíram para outro board ou cancelaram no período.',
      ['Coluna', 'Clientes que passaram', 'Tempo médio', 'Tempo médio (h)', 'Mediana', 'Maior tempo', 'Maior (h)', 'Tempo total (h)', '% do tempo total'],
      linhasG, { fmts: [null, '0', null, '0.0', null, null, '0.0', '0.0', '0.0%'], aligns: ['left', 'right', 'right', 'right', 'right', 'right', 'right', 'right', 'right'], min: 12 });
    if (linhasG.length) { escala(wsG, 'D5:D' + (4 + linhasG.length), 'FFF4A261'); escala(wsG, 'I5:I' + (4 + linhasG.length), 'FF8FB8F0'); }

    // ===== Destinos e etapas =====
    var wsD = wb.addWorksheet('Destinos e etapas', { properties: { tabColor: { argb: COR.azul } }, views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
    titulo(wsD, 'Para onde foram e em que etapa desistiram', sub, 5);
    cabecalho(wsD, 4, ['Para onde foram (outro board)', 'Clientes'], 1);
    cabecalho(wsD, 4, ['Etapa em que desistiram', 'Cancelamentos'], 4);
    var dest = d.destinos.map(function (x) { return [x.nome, x.qtd]; });
    var etap = d.etapas.map(function (x) { return [x.nome, x.qtd]; });
    if (dest.length) linhasTabela(wsD, 5, dest, [null, '0'], ['left', 'right'], 1);
    if (etap.length) linhasTabela(wsD, 5, etap, [null, '0'], ['left', 'right'], 4);
    if (!dest.length) { wsD.getCell(5, 1).value = 'Nenhum registro.'; wsD.getCell(5, 1).font = { italic: true, color: { argb: COR.cinza } }; }
    if (!etap.length) { wsD.getCell(5, 4).value = 'Nenhum registro.'; wsD.getCell(5, 4).font = { italic: true, color: { argb: COR.cinza } }; }
    wsD.getColumn(1).width = 36; wsD.getColumn(2).width = 14; wsD.getColumn(3).width = 4; wsD.getColumn(4).width = 36; wsD.getColumn(5).width = 16;
    paginaCfg(wsD);

    // ===== Clientes =====
    var cab = ['Nº do cliente', 'Nome', 'Card (nome completo)', 'Situação', 'Entrada', 'Saída', 'Tempo no board', 'Horas', 'Destino / etapa'];
    d.colunas.forEach(function (c) { cab.push('Tempo em: ' + c.nome); });
    var linhasC = d.clientes.map(function (c) {
      var l = [c.numero, c.nome, c.card, SITUACAO[c.situacao], loc(c.entrada), loc(c.saida), R.fmtTempo(c.ms), R.horas(c.ms),
               c.situacao === 'out' ? c.destino : 'Desistiu em: ' + (c.etapa || '—')];
      d.colunas.forEach(function (col) { var x = c.colunas && c.colunas[col.id]; l.push(x && x.ms > 0 ? R.fmtTempo(x.ms) : ''); });
      return l;
    });
    var fmtsC = [null, null, null, null, FMT_DATAHORA, FMT_DATAHORA, null, '0.0', null];
    var wsC = abaDados(wb, 'Clientes', COR.navy, 'Tempo de cada cliente', sub, cab, linhasC,
      { fmts: fmtsC, aligns: ['left', 'left', 'left', 'left', 'right', 'right', 'right', 'right', 'left'], congelarCols: 2, max: 40 });
    linhasC.forEach(function (l, i) {
      var cel = wsC.getCell(5 + i, 4);
      var cancel = d.clientes[i].situacao === 'cancel';
      preencher(cel, cancel ? COR.vermelhoL : COR.azulL);
      cel.font = { name: 'Calibri', size: 10, bold: true, color: { argb: cancel ? COR.vermelho : COR.azul } };
    });
    if (linhasC.length) escala(wsC, 'H5:H' + (4 + linhasC.length), 'FFF4A261');

    // ===== Horas por coluna =====
    if (d.colunas.length) {
      var cabH = ['Nº do cliente', 'Nome', 'Situação'];
      d.colunas.forEach(function (c) { cabH.push(c.nome + ' (h)'); });
      cabH.push('Soma (h)');
      var linhasH = d.clientes.map(function (c) {
        var l = [c.numero, c.nome, SITUACAO[c.situacao]], soma = 0;
        d.colunas.forEach(function (col) {
          var x = c.colunas && c.colunas[col.id], h = x && x.ms > 0 ? R.horas(x.ms) : '';
          if (h !== '') soma += h;
          l.push(h);
        });
        l.push(Math.round(soma * 10) / 10);
        return l;
      });
      var fmtsH = [null, null, null].concat(d.colunas.map(function () { return '0.0'; })).concat(['0.0']);
      var wsH = abaDados(wb, 'Horas por coluna', COR.laranja, 'Horas de cada cliente em cada coluna', sub + '   |   Quanto mais forte a cor, mais tempo.', cabH, linhasH, { fmts: fmtsH, congelarCols: 2, min: 12, max: 24 });
      if (linhasH.length) escala(wsH, 'D5:' + colLetra(3 + d.colunas.length) + (4 + linhasH.length), 'FFF4A261');
    }

    // ===== Cards (movimentações) =====
    var linhasK = d.cards.map(function (c) { return [loc(c.date), TIPO[c.tipo], c.motivo, c.card, c.origem, c.destino, c.quem, c.obs || '']; });
    var wsK = abaDados(wb, 'Cards', COR.cinza, 'Cards que entraram, saíram ou cancelaram', sub,
      ['Data', 'Tipo', 'Motivo', 'Card', 'Lista de origem', 'Lista de destino', 'Responsável', 'Observação'], linhasK,
      { fmts: [FMT_DATAHORA], aligns: ['right', 'left', 'left', 'left', 'left', 'left', 'left', 'left'], max: 60 });
    d.cards.forEach(function (c, i) {
      var cel = wsK.getCell(5 + i, 2);
      var cor = c.tipo === 'in' ? [COR.verdeL, COR.verde] : c.tipo === 'out' ? [COR.azulL, COR.azul] : [COR.vermelhoL, COR.vermelho];
      preencher(cel, cor[0]);
      cel.font = { name: 'Calibri', size: 10, bold: true, color: { argb: cor[1] } };
    });

    // ===== Pendentes =====
    if (d.pendentes.length) {
      var linhasP = d.pendentes.map(function (e) { return [loc(e.date), TIPO[e.tipo], e.card, e.motivo, e.porque, 'Sim (provisório)']; });
      var wsP = abaDados(wb, 'Pendentes', COR.laranja, 'Pendentes de confirmação', 'Já estão contados provisoriamente. Confirme no painel (aba Pendentes). ' + sub,
        ['Data', 'Evento', 'Card', 'O que foi', 'Por que precisa de confirmação', 'Está contando?'], linhasP,
        { fmts: [FMT_DATAHORA], aligns: ['right', 'left', 'left', 'left', 'left', 'left'], max: 70 });
      linhasP.forEach(function (l, i) { var c = wsP.getCell(5 + i, 6); preencher(c, 'FFFFF4E5'); c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COR.laranja } }; });
    }

    // ===== Ignorados =====
    if (d.ignorados.length) {
      var linhasI = d.ignorados.map(function (e) { return [loc(e.date), TIPO[e.tipo], e.card, e.motivo, e.porque]; });
      abaDados(wb, 'Ignorados', COR.vermelho, 'Ignorados por provável engano', 'Estes eventos não entram nas contagens. ' + sub,
        ['Data', 'Evento', 'Card', 'O que foi', 'Por que foi ignorado'], linhasI,
        { fmts: [FMT_DATAHORA], aligns: ['right', 'left', 'left', 'left', 'left'], max: 60 });
    }
    return wb;
  };

  function baixarBlob(blob, nome) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = nome;
    document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 2000);
  }

  X.baixarBonito = async function (dados, nome) {
    if (typeof root.ExcelJS === 'undefined') throw new Error('A biblioteca de Excel não carregou');
    var wb = X.construir(root.ExcelJS, dados);
    var buf = await wb.xlsx.writeBuffer();
    baixarBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), nome);
  };

  // ---------- Versão simples (reserva) ----------
  X.baixarSimples = function (d, nome) {
    var wb = XLSX.utils.book_new(), k = d.kpis;
    var res = [['Período', d.periodoTxt], ['Filtros', d.filtrosTxt], [],
      ['Entradas', k.in], ['Saídas (outro board)', k.out], ['Cancelamentos', k.cancel], ['Saldo', k.saldo], [],
      [d.serieRotulo, 'Entradas', 'Saídas', 'Cancelamentos', 'Saldo']]
      .concat(d.serie.map(function (s) { return [s.rotulo, s.in, s.out, s.cancel, s.saldo]; }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(res), 'Resumo');

    var cab = ['Nº do cliente', 'Nome', 'Card', 'Situação', 'Entrada', 'Saída', 'Tempo no board', 'Horas', 'Destino / etapa'];
    d.colunas.forEach(function (c) { cab.push('Tempo em: ' + c.nome); });
    var cli = [cab].concat(d.clientes.map(function (c) {
      var l = [c.numero, c.nome, c.card, SITUACAO[c.situacao], R.fmtFull(c.entrada), R.fmtFull(c.saida), R.fmtTempo(c.ms), R.horas(c.ms),
               c.situacao === 'out' ? c.destino : 'Desistiu em: ' + (c.etapa || '—')];
      d.colunas.forEach(function (col) { var x = c.colunas && c.colunas[col.id]; l.push(x && x.ms > 0 ? R.fmtTempo(x.ms) : ''); });
      return l;
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cli), 'Clientes');

    var cards = [['Data', 'Hora', 'Tipo', 'Motivo', 'Card', 'Lista de origem', 'Lista de destino', 'Responsável', 'Observação']]
      .concat(d.cards.map(function (c) { return [R.fmtFull(c.date), R.fmtHora(c.date), TIPO[c.tipo], c.motivo, c.card, c.origem, c.destino, c.quem, c.obs || '']; }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cards), 'Cards');

    if (d.pendentes.length) {
      var pd = [['Data', 'Evento', 'Card', 'O que foi', 'Por que precisa de confirmação']]
        .concat(d.pendentes.map(function (e) { return [R.fmtFull(e.date), TIPO[e.tipo], e.card, e.motivo, e.porque]; }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pd), 'Pendentes');
    }
    if (d.ignorados.length) {
      var ig = [['Data', 'Evento', 'Card', 'O que foi', 'Por que foi ignorado']]
        .concat(d.ignorados.map(function (e) { return [R.fmtFull(e.date), TIPO[e.tipo], e.card, e.motivo, e.porque]; }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ig), 'Ignorados');
    }
    XLSX.writeFile(wb, nome);
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = X; else root.ExcelFlow = X;
})(typeof window !== 'undefined' ? window : this);
