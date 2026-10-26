/* Consultas à API do Trello (somente leitura). */
(function (root) {
  'use strict';
  var T = {};
  var BASE = 'https://api.trello.com/1';

  T.FILTER = 'createCard,copyCard,moveCardToBoard,moveCardFromBoard,deleteCard,updateCard:closed,updateCard:idList';
  T.HIST_FILTER = 'createCard,copyCard,moveCardToBoard,moveCardFromBoard,updateCard:idList,updateCard:closed';

  function esperar(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  // fetch com nova tentativa quando o Trello pede para ir mais devagar (erro 429)
  async function buscar(url, erroMsg) {
    for (var tentativa = 0; tentativa < 4; tentativa++) {
      var r = await fetch(url);
      if (r.status === 429) { await esperar(1200 * (tentativa + 1)); continue; }
      if (!r.ok) { var err = new Error('Erro ' + r.status + ' ' + (erroMsg || 'ao consultar o Trello')); err.status = r.status; throw err; }
      return r.json();
    }
    throw new Error('O Trello pediu para aguardar (muitas consultas). Tente de novo em alguns segundos.');
  }
  function cred(token) { return '&key=' + root.APP_KEY + '&token=' + token; }

  T.fetchActions = async function (boardId, since, until, token, filter) {
    var all = [], before = until ? until.toISOString() : null;
    while (true) {
      var url = BASE + '/boards/' + boardId + '/actions?filter=' + (filter || T.FILTER) +
        '&limit=1000&since=' + since.toISOString() + cred(token) + (before ? '&before=' + before : '');
      var a = await buscar(url);
      all = all.concat(a);
      if (a.length < 1000) break;
      before = a[a.length - 1].id;
    }
    var vistos = {};
    return all.filter(function (x) { if (vistos[x.id]) return false; vistos[x.id] = true; return true; });
  };

  // Dados atuais de cada card: etiquetas, nome e descrição (inclui arquivados; cards excluídos não existem mais)
  T.fetchCards = async function (boardId, token) {
    var labels = {}, info = {}, before = null;
    while (true) {
      var url = BASE + '/boards/' + boardId + '/cards/all?fields=idLabels,name,desc&limit=1000' + cred(token) + (before ? '&before=' + before : '');
      var a = await buscar(url, 'ao buscar os dados dos cards');
      var novos = 0, menor = null;
      a.forEach(function (c) {
        if (!(c.id in labels)) {
          labels[c.id] = c.idLabels || [];
          info[c.id] = { name: c.name || '', desc: String(c.desc || '').slice(0, 1500) };
          novos++;
        }
        if (menor === null || c.id < menor) menor = c.id;
      });
      if (a.length < 1000 || novos === 0) break;
      before = menor;
    }
    return { labels: labels, info: info };
  };

  T.fetchLabels = async function (boardId, token) {
    try { return await buscar(BASE + '/boards/' + boardId + '/labels?fields=name,color&limit=1000' + cred(token)); }
    catch (e) { return []; }
  };

  // O Trello só manda o ID do outro board; aqui buscamos o nome pelo ID
  T.resolverBoards = async function (actions, token, nomes) {
    var ids = {};
    actions.forEach(function (a) {
      var d = a.data || {};
      var b = a.type === 'moveCardFromBoard' ? d.boardTarget : a.type === 'moveCardToBoard' ? d.boardSource : null;
      if (b && b.id && !b.name && !nomes[b.id]) ids[b.id] = true;
    });
    await Promise.all(Object.keys(ids).map(async function (id) {
      try {
        var j = await buscar(BASE + '/boards/' + id + '?fields=name' + cred(token));
        if (j && j.name) nomes[id] = j.name;
      } catch (e) { /* se falhar, o nome fica "undefined" */ }
    }));
  };

  root.TrelloApi = T;
})(window);
