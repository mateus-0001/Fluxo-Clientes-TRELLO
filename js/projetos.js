/* Classificacao de projetos pelas etiquetas comerciais do SETOR DE TRIAGEM.
   Um card = um atendimento; projetos podem ser 1 ou 2.
   Etiquetas operacionais, financeiras e de andamento nao classificam produtos.
   As etiquetas atuais nao provam a classificacao historica do atendimento. */
(function (root) {
  'use strict';
  var P = {};
  function norm(s) {
    return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/[^A-Z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  function tem(s, rx) { return rx.test(s); }
  function etiquetaProduto(raw) {
    var s = norm(raw);
    if (!s) return null;

    // Etiquetas de fluxo, comunicacao, pagamento, equipe e produtos adicionais
    // nunca definem sozinhas o servico principal.
    if (tem(s, /^(?:FALTAM?|PAGOU|NAO |JA |FORMS? |FORMULARIO |CONTRATO |ENCAMINHADO|AGUARDANDO|RESGATE|TENTATIVA|LOHANNA|PRIORIDADE|CONCLUIDO|COMPRAR|DESISTENTE|ADICIONAL|OPORTUNIDADE|BRINDE|FALTA |VOLTAR|EMAIL|CORREO|INFORMACAO|COLABORADOR|COMENTARIO|VALOR|DESCONTO|WHATS|LIGAR|STANDBY|IR PARA|CLIENTE SEM|CARTAO|FICHA|2 PACK\b)/)) return null;
    if (tem(s, /\b(?:ENVIADO|PREENCHIDO|RESPONDEU|COBRAR|FEITO|REALIZADA|POR NOS)\b/) &&
        !tem(s, /^COMBO\b/)) return null;

    // Regra explicita do setor: QUALQUER etiqueta com a palavra COMBO e logo + site.
    if (tem(s, /\bCOMBO\b|\bLOGO\s*(?:E|COM|MAIS)\s*(?:SITE|ECOMMERCE|E COMMERCE)\b|\bLOGO\s+SITE\b/)) return 'combo';

    // HOT POCKET e uma modalidade de logo. PACOTE BASICO / ESENCIAL / AVANZADO
    // e outras variantes comerciais do logo tambem representam um logo.
    if (tem(s, /\bHOT POCKET\b|\bVETORIZACAO\b|\bMELHORIA DE LOGO\b/)) return 'logo';
    if (tem(s, /^PACOTE\s+(?:BASICO|ESENCIAL|ESSENCIAL|AVANZADO|AVANCADO)(?:\b|$)/)) {
      // Mistura de duas familias no mesmo nome: nao presumir combo automaticamente.
      if (tem(s, /\b(?:ECOM|ECOMMERCE|SITE|SHOPIFY|WOOCOMMERCE)\b/)) return 'incerto';
      return 'logo';
    }
    if (tem(s, /\b(?:LOGO|LOGOTIPO|LOGOMARCA)\b/)) return 'logo';

    // Etiquetas de site. Inclui explicitamente os pacotes personalizados
    // e reestruturacao de Shopify, que nao contem a palavra SITE.
    if (tem(s, /\b(?:PERSONALIZADO|PERSONALIZADA|PERSONALIZADOS|PERSONALIZADAS)\b/)) return 'site';
    if (tem(s, /\b(?:REESTRUTURACAO|REESTRUCTURACION)\s+(?:DE\s+)?SHOPIFY\b/)) return 'site';
    if (tem(s, /\b(?:SITE|WEBSITE|SHOPIFY|WOOCOMMERCE|ECOM|ECOMMERCE|E COMMERCE)\b/)) return 'site';
    if (tem(s, /^(?:STARTER|DELUXE|PROFESIONAL|PROFISSIONAL|PROFESSIONAL|ADVANCED)(?:\b|$)/)) return 'site';

    // Outros produtos/etiquetas nao devem virar projetos por suposicao.
    return null;
  }

  P.classificar = function (nomes, doisProjetos) {
    var combo = false, logo = false, site = false, incerto = false;
    (nomes || []).forEach(function (nome) {
      var produto = etiquetaProduto(nome);
      if (produto === 'combo') combo = true;
      else if (produto === 'logo') logo = true;
      else if (produto === 'site') site = true;
      else if (produto === 'incerto') incerto = true;
    });
    if (combo || (logo && site && !incerto)) return { tipo:'combo', total:2, logo:1, site:1, confiavel:true };
    if (incerto) return { tipo:'nao-identificado', total:0, logo:0, site:0, confiavel:false };
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
