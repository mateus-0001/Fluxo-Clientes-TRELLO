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

    // Adicionais isolados sao acompanhados separadamente mas nao sao projetos contados.
    if (/^(?:ADICIONAL(?:\b|$)|FLYER ADICIONAL\b|OPORTUNIDADE DE ADICIONAL\b)/.test(s)) return 'adicional';
    // Outros produtos independentes, diferentes de adicionais.
    if (/^(?:ARTE DE EMBALAGEM|PACK DE POSTS|PACK DE [0-9]+ POSTS?|CARTAO DE VISITAS?|CARTAO FIDELIDAD|SO VETORIZACAO|PLUGIN DE CHATBOT NA WEB)(?:\b|$)/.test(s)) return 'outro';
    // E-mail profissional e correio profissional sao ADICIONAIS, nunca projetos.
    if (/^(?:EMAIL PROFISSIONAL|CORREO PROFESIONAL|CORREO PROFISSIONAL|E MAIL PROFISSIONAL)(?:\b|$)/.test(s)) return 'adicional';

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
    if (tem(s, /^ELDER$/)) return 'site';
    if (tem(s, /\b(?:PERSONALIZADO|PERSONALIZADA|PERSONALIZADOS|PERSONALIZADAS)\b/)) return 'site';
    if (tem(s, /\b(?:REESTRUTURACAO|REESTRUCTURACION)\s+(?:DE\s+)?SHOPIFY\b/)) return 'ecommerce';
    if (tem(s, /\b(?:SHOPIFY|WOOCOMMERCE|ECOM|ECOMMERCE|E COMMERCE)\b/)) return 'ecommerce';
    if (tem(s, /\b(?:SITE|WEBSITE)\b/)) return 'site';
    if (tem(s, /^(?:STARTER|DELUXE|PROFESIONAL|PROFISSIONAL|PROFESSIONAL|ADVANCED)(?:\b|$)/)) return 'site';

    // Outros produtos/etiquetas nao devem virar projetos por suposicao.
    return null;
  }

  // Uma etiqueta de COMBO significa exatamente 2 projetos (1 logo + 1 site).
  // As etiquetas classificam o atendimento atual, mas nao provam o tipo historico.
  P.classificar = function (nomes, doisProjetos) {
    var grupos = { combo: [], logo: [], site: [], ecommerce: [], outro: [], adicional: [], incerto: [] };
    var vistos = {};
    (nomes || []).forEach(function (original) {
      var nome = String(original || '').trim();
      if (!nome || vistos[norm(nome)]) return;
      vistos[norm(nome)] = true;
      var classe = etiquetaProduto(nome);
      if (classe && grupos[classe]) grupos[classe].push(nome);
    });
    var componentes = [];
    function adicionar(tipo, tag) { componentes.push({ categoria: tipo, etiqueta: tag }); }
    var combos = grupos.combo;
    if (combos.length) {
      adicionar('Logo', combos[0]);
      adicionar(/\b(?:ECOM|ECOMMERCE|SHOPIFY|WOOCOMMERCE)\b/.test(norm(combos[0])) ? 'E-commerce' : 'Site', combos[0]);
    } else if (grupos.incerto.length) {
      return { tipo: 'nao-identificado', total: 0, logo: 0, site: 0, ecommerce: 0, outros: 0,
        confiavel: false, componentes: [], etiquetas: nomes || [] };
    } else {
      if (grupos.logo.length) adicionar('Logo', grupos.logo[0]);
      // Uma etiqueta e-commerce tem prioridade sobre uma etiqueta generica de site.
      if (grupos.ecommerce.length) adicionar('E-commerce', grupos.ecommerce[0]);
      else if (grupos.site.length) adicionar('Site', grupos.site[0]);
      // Etiquetas de adicionais explicitamente identificados como produto
      // podem representar outro projeto no mesmo atendimento.
      if (grupos.outro.length) adicionar('Outro', grupos.outro[0]);
    }
    if (!componentes.length && grupos.adicional.length) {
      return { tipo: 'adicional-desconsiderado', total: 0, logo: 0, site: 0, ecommerce: 0, outros: 0,
        confiavel: true, componentes: [], etiquetas: nomes || [] };
    }
    if (!componentes.length) {
      return { tipo: doisProjetos ? 'dois-nao-identificados' : 'nao-identificado',
        total: 0, logo: 0, site: 0, ecommerce: 0, outros: 0, confiavel: false,
        componentes: [], etiquetas: nomes || [] };
    }
    var nLogo = componentes.filter(function (c) { return c.categoria === 'Logo'; }).length;
    var nSite = componentes.filter(function (c) { return c.categoria === 'Site'; }).length;
    var nEcommerce = componentes.filter(function (c) { return c.categoria === 'E-commerce'; }).length;
    var nOutro = componentes.filter(function (c) { return c.categoria === 'Outro'; }).length;
    return { tipo: nLogo && (nSite || nEcommerce) ? 'combo' : nLogo && nOutro ? 'logo+outro' :
      (nSite || nEcommerce) && nOutro ? 'site+outro' : nLogo ? 'logo' : nEcommerce ? 'ecommerce' : nSite ? 'site' : 'outro',
      total: componentes.length, logo: nLogo, site: nSite, ecommerce: nEcommerce, outros: nOutro,
      confiavel: true, componentes: componentes, etiquetas: nomes || [] };
  };

  P.resumir = function (eventos, cardLabels, labels, cardInfo, R, nomesExternos) {
    var nomes = {}, detalhe = [], numeros = {}, semNumero = 0;
    (labels || []).forEach(function (l) { nomes[l.id] = l.name || ''; });
    var total = {in:0,out:0,cancel:0}, logo = {in:0,out:0,cancel:0},
      site = {in:0,out:0,cancel:0}, ecommerce = {in:0,out:0,cancel:0}, outros = {in:0,out:0,cancel:0};
    var naoIdentificados = {in:0,out:0,cancel:0}, doisSemTipo = {in:0,out:0,cancel:0};
    // Contagem informativa: SOMENTE cards com etiqueta contendo a palavra COMBO.
    // Nao presumir combo por existir logo + site sem essa etiqueta.
    var combosEtiqueta = {in:0,out:0,cancel:0};
    var agrupados = {}, eventosPorTag = {};
    (eventos || []).forEach(function (ev) {
      if (ev.ign || !Object.prototype.hasOwnProperty.call(total,ev.tipo)) return;
      var ls = (cardLabels && cardLabels[ev.id] || []).map(function (id) { return nomes[id] || ''; });
      var externos = nomesExternos && nomesExternos[ev.id] || [];
      ls = ls.concat(externos).filter(Boolean);
      var info = cardInfo && cardInfo[ev.id];
      var cl = P.classificar(ls, ev.doisProjetos);
      var temTagCombo = ls.some(function (tag) { return /\bCOMBO\b/.test(norm(tag)); });
      if (temTagCombo) combosEtiqueta[ev.tipo]++;
      var num = R.partirNome((ev.card && ev.card.name) || (info && info.name) || '').numero;
      if (ev.tipo === 'in') { if (num) numeros[num] = 1; else semNumero++; }
      if (!cl.confiavel) {
        naoIdentificados[ev.tipo]++;
        if (cl.tipo === 'dois-nao-identificados') doisSemTipo[ev.tipo]++;
      } else {
        total[ev.tipo] += cl.total;
        logo[ev.tipo] += cl.logo;
        site[ev.tipo] += cl.site;
        ecommerce[ev.tipo] += cl.ecommerce;
        outros[ev.tipo] += cl.outros;
        // Cada projeto entra UMA vez no resumo por etiqueta. Em combos,
        // a mesma etiqueta gera duas linhas: Logo e Site.
        cl.componentes.forEach(function (comp) {
          var chave = comp.categoria + '|' + comp.etiqueta;
          var row = agrupados[chave];
          if (!row) {
            row = agrupados[chave] = {categoria:comp.categoria,etiqueta:comp.etiqueta,
              in:0,out:0,cancel:0};
          }
          row[ev.tipo]++;
        });
      }
      detalhe.push({card:(ev.card && ev.card.name)||'',tipoEvento:ev.tipo,data:ev.when,
        classe:cl.tipo,projetosConfirmados:cl.confiavel?cl.total:0,
        logo:cl.logo,site:cl.site,ecommerce:cl.ecommerce,outros:cl.outros,etiquetas:ls,
        componentes:cl.componentes,temTagCombo:temTagCombo,precisaClassificar:!cl.confiavel,cardId:ev.id,
        actionId:(ev.a && ev.a.id) || ''});
    });
    var porEtiqueta = Object.keys(agrupados).map(function (k) {return agrupados[k];});
    porEtiqueta.sort(function (a,b) {
      return b.out-a.out || b.in-a.in || a.categoria.localeCompare(b.categoria) || a.etiqueta.localeCompare(b.etiqueta);
    });
    return {atendimentos:R.contar(eventos),clientesUnicos:Object.keys(numeros).length,
      semIdentificador:semNumero,projetosConfirmados:total,logo:logo,site:site,ecommerce:ecommerce,outros:outros,
      semClassificacao:naoIdentificados,doisSemTipo:doisSemTipo,combosEtiqueta:combosEtiqueta,
      detalhe:detalhe,porEtiqueta:porEtiqueta,
      nota:'Projetos classificados pelas etiquetas comerciais atualmente acessiveis. Etiquetas atuais nao provam a classificacao historica.'};
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = P;
  else root.ProjetosFlow = P;
})(typeof window !== 'undefined' ? window : this);
