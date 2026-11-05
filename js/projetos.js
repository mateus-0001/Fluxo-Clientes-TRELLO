/* Classificacao de projetos pelas etiquetas comerciais do SETOR DE TRIAGEM.
   Contagem de atendimentos permanece separada. Etiquetas financeiras e de fluxo
   nao definem o tipo de projeto. Dados de etiquetas atuais nao sao historicos. */
(function (root) {
  'use strict';
  var P = {};
  function norm(s) {
    return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/[^A-Z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  function possui(s, exp) { return exp.test(' ' + s + ' '); }
  P.classificar = function (nomes, doisProjetos) {
    var nomesNormalizados = (nomes || []).map(norm).filter(Boolean);
    var combo = false, logo = false, site = false;
    nomesNormalizados.forEach(function (s) {
      // Combo prevalece: um atendimento, dois projetos (logo + site).
      if (possui(s, /\bCOMBO\b|\bLOGO\s+(?:E|COM)\s+(?:SITE|ECOMMERCE)\b|\bLOGO\s+SITE\b|\bSITE\s+LOGO\b/)) combo = true;
      // Produtos de identidade visual (grafia em portugues e espanhol).
      if (possui(s, /\bLOGO(?:TIPO|MARCA)?\b|\bPACOTE\s+(?:BASICO|ESENCIAL|ESSENCIAL|AVANZADO|AVANCADO)\b|\bVETORIZACAO\b/)) logo = true;
      // Produtos de sites e ecommerce, com e sem o prefixo SITE.
      if (possui(s, /\bSITE\b|\bWEBSITE\b|\bSHOPIFY\b|\bWOOCOMMERCE\b|\bECOM\b|\bE\s*COMMERCE\b|\bSTARTER\b|\bDELUXE\b|\bPROFESIONAL\b|\bPROFISSIONAL\b|\bPROFESSIONAL\b/)) site = true;
      // ADVANCED isolado e incerto; ECOM ADVANCED ja e site.
    });
    if (combo || (logo && site)) return { tipo:'combo', total:2, logo:1, site:1, confiavel:true };
    if (doisProjetos) return { tipo:'dois-nao-identificados', total:2, logo:0, site:0, confiavel:false };
    if (logo) return { tipo:'logo', total:1, logo:1, site:0, confiavel:true };
    if (site) return { tipo:'site', total:1, logo:0, site:1, confiavel:true };
    return { tipo:'nao-identificado', total:0, logo:0, site:0, confiavel:false };
  };
  P.resumir = function (eventos, cardLabels, labels, cardInfo, R, nomesExternos) {
    var nomes = {}, detalhe = [], numeros = {}, semNumero = 0;
    (labels || []).forEach(function (l) { nomes[l.id] = l.name || ''; });
    var total = {in:0,out:0,cancel:0}, logo = {in:0,out:0,cancel:0}, site = {in:0,out:0,cancel:0};
    var naoIdentificados = {in:0,out:0,cancel:0}, doisSemTipo = {in:0,out:0,cancel:0};
    (eventos || []).forEach(function (ev) {
      if (ev.ign || !Object.prototype.hasOwnProperty.call(total,ev.tipo)) return;
      var ls = (cardLabels && cardLabels[ev.id] || []).map(function (id) { return nomes[id] || ''; });
      var externos = nomesExternos && nomesExternos[ev.id] || [];
      ls = ls.concat(externos);
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
      detalhe.push({card:(ev.card && ev.card.name)||'',tipoEvento:ev.tipo,data:ev.when,
        classe:cl.tipo,projetosConfirmados:cl.confiavel?cl.total:0,
        logo:cl.logo,site:cl.site,precisaClassificar:!cl.confiavel,cardId:ev.id});
    });
    return {atendimentos:R.contar(eventos),clientesUnicos:Object.keys(numeros).length,
      semIdentificador:semNumero,projetosConfirmados:total,logo:logo,site:site,
      semClassificacao:naoIdentificados,doisSemTipo:doisSemTipo,detalhe:detalhe,
      nota:'Tipo de projeto inferido de etiquetas de produto (inclusive nos quadros de destino acessiveis). Etiquetas atuais nao garantem classificacao historica.'};
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = P;
  else root.ProjetosFlow = P;
})(typeof window !== 'undefined' ? window : this);
