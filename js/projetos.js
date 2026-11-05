/* Indicadores adicionais, independentes da contagem oficial de eventos.
   Tipos são inferidos APENAS de etiquetas explícitas. Não presume que pacote = logo. */
(function (root) {
  'use strict';
  var P = {};
  function norm(s) {
    return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
  }
  P.classificar = function (nomes, doisProjetos) {
    var l = (nomes || []).map(norm);
    var combo = l.some(function (s) { return /\b(COMBO|LOGO\s*\+\s*SITE|SITE\s*\+\s*LOGO|LOGO\s+E\s+SITE)\b/.test(s); });
    var logo = l.some(function (s) { return /\b(LOGO|LOGOTIPO|LOGOMARCA)\b/.test(s); });
    var site = l.some(function (s) { return /\b(SITE|WEBSITE|SHOPIFY|ECOM|E-COMMERCE)\b/.test(s); });
    if (combo || (logo && site)) return {tipo:'combo',total:2,logo:1,site:1,confiavel:true};
    if (doisProjetos) return {tipo:'dois-nao-identificados',total:2,logo:0,site:0,confiavel:false};
    if (logo) return {tipo:'logo',total:1,logo:1,site:0,confiavel:true};
    if (site) return {tipo:'site',total:1,logo:0,site:1,confiavel:true};
    return {tipo:'nao-identificado',total:0,logo:0,site:0,confiavel:false};
  };
  P.resumir = function (eventos, cardLabels, labels, cardInfo, R) {
    var nomes = {}, detalhe = [], numeros = {}, semNumero = 0;
    (labels || []).forEach(function (l) { nomes[l.id] = l.name || ''; });
    var total = {in:0,out:0,cancel:0}, logo = {in:0,out:0,cancel:0}, site = {in:0,out:0,cancel:0};
    var naoIdentificados = {in:0,out:0,cancel:0}, doisSemTipo = {in:0,out:0,cancel:0};
    eventos.forEach(function (ev) {
      if (ev.ign || !Object.prototype.hasOwnProperty.call(total,ev.tipo)) return;
      var ls = (cardLabels && cardLabels[ev.id] || []).map(function (id) {return nomes[id] || '';});
      var info = cardInfo && cardInfo[ev.id];
      var cl = P.classificar(ls, ev.doisProjetos);
      var num = R.partirNome((ev.card && ev.card.name) || (info && info.name) || '').numero;
      if (ev.tipo === 'in') { if (num) numeros[num] = 1; else semNumero++; }
      if (!cl.confiavel) {
        naoIdentificados[ev.tipo]++;
        if (cl.tipo === 'dois-nao-identificados') doisSemTipo[ev.tipo]++;
      } else {
        total[ev.tipo] += cl.total; logo[ev.tipo] += cl.logo; site[ev.tipo] += cl.site;
      }
      detalhe.push({card:(ev.card && ev.card.name)||'',tipoEvento:ev.tipo, data:ev.when,
        classe:cl.tipo,projetosConfirmados:cl.confiavel?cl.total:0,
        logo:cl.logo,site:cl.site,precisaClassificar:!cl.confiavel,cardId:ev.id});
    });
    return {atendimentos: R.contar(eventos),clientesUnicos:Object.keys(numeros).length,
      semIdentificador:semNumero,projetosConfirmados:total,logo:logo,site:site,
      semClassificacao:naoIdentificados,doisSemTipo:doisSemTipo,detalhe:detalhe,
      nota:'Projetos confirmados por etiqueta explícita. Cards sem tipo permanecem separados. Etiquetas atuais podem faltar em cards transferidos.'};
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = P; else root.ProjetosFlow = P;
})(typeof window !== 'undefined' ? window : this);
