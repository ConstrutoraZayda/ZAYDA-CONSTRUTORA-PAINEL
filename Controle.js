/**
 * ZAYDA CONSTRUTORA | CONTROLE DE MATERIAIS POR OBRA
 * Projeto Apps Script VINCULADO à planilha "ZAYDA CONSTRUTORA | CONTROLE DE MATERIAIS POR OBRA"
 * (abrir essa planilha > Extensões > Apps Script). NÃO colar no projeto do formulário.
 *
 * Arquivos deste projeto: Controle.gs (este) e Painel.html.
 *
 * O que este script faz:
 *  1. Sincronização: copia PEDIDOS, ITENS_PEDIDO e CATALOGO da planilha do formulário
 *     para as abas ocultas _IMPORT_PEDIDOS, _IMPORT_ITENS e _IMPORT_CATALOGO.
 *     Pedidos cancelados ou encerrados (aba CANCELAMENTOS do formulário) vão para _IMPORT_CANCELAMENTOS.
 *     Cada entrega (aba ENTREGAS, com quem entregou e quem recebeu) vai para _IMPORT_ENTREGAS.
 *     É uma cópia completa a cada execução (idempotente): rodar duas vezes não duplica nada.
 *     A planilha do formulário só é LIDA, nunca alterada.
 *     Também copia a aba OBRAS da planilha "Controle e Planejamento de Obras" (ID em
 *     CONFIGURACOES!B13) para a aba OBRAS deste controle, colunas A a J. Essa planilha
 *     também é só LIDA. As obras são cadastradas e editadas lá, não aqui.
 *     E as entregas confirmadas no formulário (aba ENTREGAS da planilha do formulário): cada uma
 *     vira uma linha em MATERIAIS_E_CUSTOS já CONFIRMADA, com o custo = preço unitário do pedido
 *     (decisão da Zayda em 29/10/2026). Item sem preço no pedido fica PENDENTE DE CUSTO. Só acrescenta
 *     linhas novas (pelo ID ENT-...); nunca altera nem apaga linhas já existentes.
 *  2. Painel executivo (app da Web), somente leitura.
 *  3. Menu "Zayda" na planilha.
 *
 * Instalação: ver o guia de instalação entregue junto com este arquivo.
 */
const CTRL = Object.freeze({
  fuso: 'America/Sao_Paulo',
  prefixoPedido: 'ZAYDA-',
  conf: 'CONFIGURACOES',
  celulaModo: 'B5', celulaOrigemId: 'B6', celulaAbaPedidos: 'B7', celulaAbaItens: 'B8',
  celulaAbaCatalogo: 'B9', celulaUltimaSinc: 'B10', celulaObrasId: 'B13', celulaIgnorados: 'B14',
  abaObrasOrigem: 'OBRAS', padraoIdObra: /^OB-\d+$/i, linhaObras: 5,
  abaEntregasOrigem: 'ENTREGAS', prefixoEntrega: 'ENT-', linhaLanc: 5, linhaApelidos: 18,
  destPedidos: '_IMPORT_PEDIDOS', destItens: '_IMPORT_ITENS', destCatalogo: '_IMPORT_CATALOGO',
  obras: 'OBRAS', solicitacoes: 'SOLICITACOES', lancamentos: 'MATERIAIS_E_CUSTOS',
  abaEtapasOrigem: 'ETAPAS', destEtapas: '_IMPORT_ETAPAS', orcamentos: 'ORCAMENTOS',
  abaCancelOrigem: 'CANCELAMENTOS', destCancel: '_IMPORT_CANCELAMENTOS', destEntregas: '_IMPORT_ENTREGAS',
  intervaloMinutos: 5, copiaCompletaMinutos: 15
});

/* ============================== MENU ============================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Zayda')
    .addItem('Sincronizar formulário agora', 'sincronizarComAviso')
    .addItem('Ativar sincronização automática (formulário e obras)', 'instalarSincronizacao')
    .addSeparator()
    .addItem('Abrir painel executivo', 'mostrarLinkPainel')
    .addItem('Código de acesso do painel', 'definirCodigoPainel')
    .addItem('Criar/atualizar aba de orçamentos', 'criarAbaOrcamentos')
    .addToUi();
}

/* ========================== SINCRONIZAÇÃO ========================= */
/** A planilha de controle. No acionador "ao editar" da planilha de obras, a planilha ativa é a de obras;
 *  por isso o ID do controle fica guardado (gravado em Ativar sincronização automática e a cada uso normal). */
function planilhaControle_() {
  const props = PropertiesService.getScriptProperties();
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa && ativa.getSheetByName(CTRL.conf)) {
    if (props.getProperty('ID_CONTROLE') !== ativa.getId()) props.setProperty('ID_CONTROLE', ativa.getId());
    return ativa;
  }
  const id = props.getProperty('ID_CONTROLE');
  if (!id) throw new Error('Abra a planilha de controle e use Zayda > Ativar sincronização automática.');
  return SpreadsheetApp.openById(id);
}
function texto_(v) { return String(v === null || v === undefined ? '' : v).trim(); }
/** Texto digitado no formulário que a planilha possa ter virado hora ou data ("09:00", "10/11"). */
function textoForm_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, CTRL.fuso, 'yyyy-MM-dd') === '1899-12-30'
    ? Utilities.formatDate(v, CTRL.fuso, 'HH:mm') : Utilities.formatDate(v, CTRL.fuso, 'dd/MM/yyyy');
  return texto_(v);
}

/** Número a partir de número ou texto ("R$ 1.234,56", "17,00"). Vazio ou inválido -> ''. */
function valor_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : '';
  let s = texto_(v).replace(/R\$/g, '').replace(/\s/g, '');
  if (!s) return '';
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : '';
}

/** Converte Date, texto ISO ou texto no formato Date.toString() para "yyyy-MM-dd HH:mm:ss" em São Paulo. */
function dataHora_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, CTRL.fuso, 'yyyy-MM-dd HH:mm:ss');
  const s = texto_(v);
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.length >= 19 ? s.slice(0, 19) : s.slice(0, 10);
  const d = new Date(s); // cobre "Wed Oct 28 2026 13:29:10 GMT-0300 (...)"
  return isNaN(d) ? s : Utilities.formatDate(d, CTRL.fuso, 'yyyy-MM-dd HH:mm:ss');
}

function lerOrigem_() {
  const conf = planilhaControle_().getSheetByName(CTRL.conf);
  const id = texto_(conf.getRange(CTRL.celulaOrigemId).getValue());
  if (!id) throw new Error('Preencha CONFIGURACOES!' + CTRL.celulaOrigemId + ' com o ID da planilha CATÁLOGO E SOLICITAÇÕES.');
  const origem = SpreadsheetApp.openById(id);
  const aba = cel => {
    const nome = texto_(conf.getRange(cel).getValue());
    const sh = origem.getSheetByName(nome);
    if (!sh) throw new Error('Aba "' + nome + '" não encontrada na planilha de origem.');
    return sh;
  };
  const valores = (sh, nCols) => sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), nCols).getValues() : [];

  // Pedidos de teste do formulário que não devem entrar no controle (lista em CONFIGURACOES).
  const ignorar = texto_(conf.getRange(CTRL.celulaIgnorados).getValue()).toUpperCase()
    .split(/[\s,;]+/).filter(Boolean);
  const valido = r => texto_(r[0]).indexOf(CTRL.prefixoPedido) === 0 && ignorar.indexOf(texto_(r[0]).toUpperCase()) < 0;
  // Filtra pelo prefixo do ID: funciona com ou sem cabeçalho na linha 1 da origem.
  // 15ª coluna (O) = ponto de referência, gravada pelo formulário 1.4. É lida à parte para não mudar
  // as 14 colunas originais nem a coluna O de _IMPORT_PEDIDOS (obra identificada).
  // [Função 6] ETAPA: procurada pelo nome na 1ª linha; a aba PEDIDOS do formulário não tem cabeçalho, então, sem o nome, vale a coluna P
  // (16ª, onde o formulário 3.1 grava). Coluna ausente ou vazia: etapa vazia ("Sem etapa informada" no painel).
  const shPed = aba(CTRL.celulaAbaPedidos), nColP = Math.min(Math.max(shPed.getLastColumn(), 1), 40);
  const cabPed = shPed.getLastRow() ? shPed.getRange(1, 1, 1, nColP).getValues()[0].map(v => texto_(v).toUpperCase()) : [];
  const iEtapaP = cabPed.indexOf('ETAPA') >= 0 ? cabPed.indexOf('ETAPA') : 15;
  const pedidos = valores(shPed, Math.min(Math.max(15, iEtapaP + 1), shPed.getMaxColumns()))
    .filter(valido)
    .map(r => { const p = [texto_(r[0]), Number(r[1]) || 0, dataHora_(r[2]), dataHora_(r[3]), texto_(r[4]), texto_(r[5]),
               texto_(r[6]), texto_(r[7]), texto_(r[8]), texto_(r[9]), valor_(r[10]),
               Number(valor_(r[11])) || 0, texto_(r[12]), texto_(r[13])]; p.referencia = texto_(r[14]); p.etapa = textoForm_(r[iEtapaP]); return p; });
  const itens = valores(aba(CTRL.celulaAbaItens), 11)
    .filter(valido)
    .map(r => [texto_(r[0]), Number(r[1]) || 0, texto_(r[2]), texto_(r[3]), texto_(r[4]), texto_(r[5]),
               Number(valor_(r[6])) || 0, valor_(r[7]), valor_(r[8]), texto_(r[9]), dataHora_(r[10])]);
  // Entregas confirmadas no formulário (aba criada pelo formulário 1.2; pode ainda não existir).
  const shEnt = origem.getSheetByName(CTRL.abaEntregasOrigem);
  // Coluna N (ENTREGUE_POR, formulário 2.0): quem levou o material; nas entregas antigas fica vazia.
  const entregas = (shEnt ? valores(shEnt, Math.min(14, shEnt.getMaxColumns())) : [])
    .filter(r => texto_(r[0]).indexOf(CTRL.prefixoEntrega) === 0 && valido([r[1]]))
    .map(r => ({id: texto_(r[0]), pedido: texto_(r[1]), sku: texto_(r[3]), qtd: Number(valor_(r[6])) || 0,
                data: dataHora_(r[7]).slice(0, 10), email: texto_(r[8]), obs: texto_(r[10]), recebido: texto_(r[11]),
                produto: texto_(r[4]), un: texto_(r[5]), recebidoTxt: textoForm_(r[11]), obsTxt: textoForm_(r[10]), entreguePor: textoForm_(r[13])}));
  // Pedidos cancelados ou encerrados (aba criada pelo formulário 1.9 no 1º cancelamento; pode ainda não existir).
  // A PEDIDO_ID · B TIPO · C DATA_HORA · D EMAIL_LOGIN · E MOTIVO · F ITENS_NAO_ENTREGUES
  const shCanc = origem.getSheetByName(CTRL.abaCancelOrigem);
  const cancelamentos = (shCanc ? valores(shCanc, 6) : []).filter(valido)
    .map(r => [texto_(r[0]), texto_(r[1]).toUpperCase(), dataHora_(r[2]), texto_(r[3]), textoForm_(r[4]), textoForm_(r[5])]);
  const catalogo = valores(aba(CTRL.celulaAbaCatalogo), 6)
    .filter(r => texto_(r[0]) && texto_(r[0]).toUpperCase() !== 'SKU')
    .map(r => [texto_(r[0]), texto_(r[1]), texto_(r[2]), texto_(r[3]), texto_(r[4]),
               valor_(r[5])]);

  // CRIADO_EM é a data de criação do pedido e não muda entre revisões. Usa sempre a da menor
  // revisão: corrige as revisões antigas gravadas como texto "Wed Oct 28 2026 13:29:10 GMT-0300",
  // que ficaram 4h adiantadas por diferença de fuso entre o script e a planilha do formulário.
  const criacao = {};
  pedidos.forEach(p => { const c = criacao[p[0]]; if (p[2] && (!c || p[1] < c.rev)) criacao[p[0]] = {rev: p[1], data: p[2]}; });
  pedidos.forEach(p => { if (criacao[p[0]]) p[2] = criacao[p[0]].data; });

  // Conferências de integridade (não bloqueiam, ficam no resumo).
  const chaves = {}, dup = [];
  pedidos.forEach(p => { const k = p[0] + '|R' + p[1]; if (chaves[k]) dup.push(k); chaves[k] = true; });
  const orfaos = uniq_(itens.map(i => i[0] + '|R' + i[1]).filter(k => !chaves[k]));
  // Itens sem a linha correspondente em PEDIDOS (gravação interrompida) não são importados:
  // sem isso eles seriam tratados como revisão vigente e somariam em dobro.
  const itensValidos = itens.filter(i => chaves[i[0] + '|R' + i[1]]);
  return {pedidos, itens: itensValidos, catalogo, dup, orfaos, entregas, cancelamentos};
}
/** Lê a aba OBRAS da planilha "Controle e Planejamento de Obras" (1 linha por casa, ID OB-xx).
 *  Origem: A ID · B CLIENTE · C ENDEREÇO · D PROTOCOLO · E DATA CONTRATO · F PRAZO · G ENTREGA PREVISTA ·
 *          H FASE ATUAL · I STATUS · J SITUAÇÃO PRAZO · K ÚLTIMA VISTORIA · L OBSERVAÇÕES.
 *  Destino (aba OBRAS, A:J): ID · CLIENTE · ENDEREÇO · NOME USADO NO FORMULÁRIO (= ID) · FASE ATUAL ·
 *          SITUAÇÃO DO PRAZO · ENTREGA PREVISTA · STATUS · DATA DO CONTRATO · (J vazia).
 *  Protocolo, última vistoria, observações e a aba DOCUMENTAÇÃO não são trazidos. */
function lerObras_() {
  const conf = planilhaControle_().getSheetByName(CTRL.conf);
  const id = texto_(conf.getRange(CTRL.celulaObrasId).getValue());
  if (!id) return {ativo: false, linhas: []};
  const planilha = SpreadsheetApp.openById(id);
  const sh = planilha.getSheetByName(CTRL.abaObrasOrigem);
  if (!sh) throw new Error('Aba "' + CTRL.abaObrasOrigem + '" não encontrada na planilha de obras.');
  const n = sh.getLastRow();
  const vals = n ? sh.getRange(1, 1, n, 12).getValues() : [];
  const vistos = {};
  const linhas = vals
    .filter(r => CTRL.padraoIdObra.test(texto_(r[0])))
    .filter(r => { const k = texto_(r[0]).toUpperCase(); if (vistos[k]) return false; vistos[k] = true; return true; })
    .map(r => {
      const idObra = texto_(r[0]).toUpperCase();
      const data = v => (v instanceof Date && !isNaN(v)) ? v : texto_(v);
      return [idObra, texto_(r[1]), texto_(r[2]), idObra, texto_(r[7]), texto_(r[9]), data(r[6]),
              texto_(r[8]), data(r[4]), ''];
    });
  const etapas = lerEtapas_(planilha);
  const pct = {}; etapas.linhas.forEach(e => { pct[e[0]] = e[e.length - 1]; });
  linhas.forEach(l => { if (l[0] in pct) l[9] = pct[l[0]]; }); // coluna J: % concluído (aba ETAPAS)
  const assinatura = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    JSON.stringify([linhas, etapas])));
  return {ativo: true, linhas, etapas, assinatura};
}

/** Aba ETAPAS da planilha de obras: ID · CLIENTE · uma coluna por etapa (Projeto ... Entrega) · % CONCLUÍDO.
 *  Devolve {cab: [nomes das etapas], linhas: [[id, cliente, status..., pct]]}. Sem a aba, listas vazias. */
function lerEtapas_(planilha) {
  const sh = planilha.getSheetByName(CTRL.abaEtapasOrigem);
  if (!sh || sh.getLastRow() < 2) return {cab: [], linhas: []};
  const vals = sh.getRange(1, 1, sh.getLastRow(), Math.min(sh.getLastColumn(), 30)).getValues();
  const iCab = vals.findIndex(r => texto_(r[0]).toUpperCase() === 'ID');
  if (iCab < 0) return {cab: [], linhas: []};
  const cabTodo = vals[iCab].map(texto_);
  const iPct = cabTodo.findIndex(c => /%/.test(c));
  const fim = iPct > 2 ? iPct : cabTodo.length;
  const cab = cabTodo.slice(2, fim).filter(Boolean);
  const vistos = {};
  const linhas = vals.slice(iCab + 1)
    .filter(r => CTRL.padraoIdObra.test(texto_(r[0])))
    .filter(r => { const k = texto_(r[0]).toUpperCase(); if (vistos[k]) return false; vistos[k] = true; return true; })
    .map(r => {
      const p = iPct > 2 ? Number(r[iPct]) : NaN;
      return [texto_(r[0]).toUpperCase(), texto_(r[1])].concat(cab.map((c, i) => texto_(r[2 + i])))
        .concat([isFinite(p) ? Math.round(p * 1000) / 1000 : '']);
    });
  return {cab, linhas};
}

/** Grava as etapas na aba oculta _IMPORT_ETAPAS (cria se não existir). */
function escreverEtapas_(ss, etapas) {
  if (!etapas || !etapas.cab.length) return;
  let sh = ss.getSheetByName(CTRL.destEtapas);
  if (!sh) { sh = ss.insertSheet(CTRL.destEtapas, ss.getNumSheets()); sh.hideSheet(); }
  sh.clearContents();
  const cab = ['ID_OBRA', 'CLIENTE'].concat(etapas.cab).concat(['% CONCLUÍDO']);
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  if (etapas.linhas.length) {
    sh.getRange(2, 1, etapas.linhas.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, etapas.linhas.length, cab.length).setValues(etapas.linhas.map(l => l.slice(0, cab.length)));
    sh.getRange(2, cab.length, etapas.linhas.length, 1).setNumberFormat('0%');
  }
}

/** Atualiza obras e etapas se a planilha de obras mudou (compara a assinatura dos dados).
 *  Chamado a cada 5 minutos, a cada consulta do painel e na hora em que alguém edita a planilha de obras. */
function atualizarObrasSeMudou_(forcar) {
  const props = PropertiesService.getScriptProperties();
  let obras;
  try { obras = lerObras_(); } catch (e) { Logger.log('Obras não lidas: ' + e.message); return false; }
  if (!obras.ativo) return false;
  if (!forcar && obras.assinatura === props.getProperty('HASH_OBRAS')) return false;
  const ss = planilhaControle_();
  const atuais = Math.max(0, ss.getSheetByName(CTRL.obras).getLastRow() - CTRL.linhaObras + 1);
  if (!obras.linhas.length && atuais > 0) return false; // planilha de obras vazia por engano: não apaga nada
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return false;
  try {
    escreverObras_(ss, obras.linhas);
    escreverEtapas_(ss, obras.etapas);
    try { escreverObrasFormulario_(obras.linhas); } catch (e) { Logger.log('Lista do formulário: ' + e.message); }
    SpreadsheetApp.flush();
    props.setProperty('HASH_OBRAS', obras.assinatura);
    return true;
  } finally { lock.releaseLock(); }
}

/** Acionador "ao editar" da planilha Controle e Planejamento de Obras (criado em Ativar sincronização automática). */
function aoEditarObras(e) {
  const aba = e && e.range ? e.range.getSheet().getName() : '';
  if (aba && aba !== CTRL.abaObrasOrigem && aba !== CTRL.abaEtapasOrigem) return;
  atualizarObrasSeMudou_(false);
}

function escreverObras_(ss, linhas) {
  const sh = ss.getSheetByName(CTRL.obras);
  const max = sh.getMaxRows();
  sh.getRange(CTRL.linhaObras, 1, max - CTRL.linhaObras + 1, 10).clearContent();
  if (!linhas.length) return;
  if (max < CTRL.linhaObras + linhas.length) sh.insertRowsAfter(max, CTRL.linhaObras + linhas.length - max);
  sh.getRange(CTRL.linhaObras, 1, linhas.length, 1).setNumberFormat('@');
  sh.getRange(CTRL.linhaObras, 4, linhas.length, 1).setNumberFormat('@');
  sh.getRange(CTRL.linhaObras, 1, linhas.length, 10).setValues(linhas);
  sh.getRange(CTRL.linhaObras, 10, linhas.length, 1).setNumberFormat('0%');
  if (texto_(sh.getRange(CTRL.linhaObras - 1, 10).getValue()) !== '% CONCLUÍDO') sh.getRange(CTRL.linhaObras - 1, 10).setValue('% CONCLUÍDO');
  // Uma falha na aba de orçamentos não pode impedir a atualização das obras.
  try { sincronizarOrcamentos_(ss, linhas); } catch (e) { Logger.log('Aba de orçamentos não atualizada: ' + e.message); }
}

/* ===================== ORÇAMENTO DE MATERIAIS ===================== */
/** Aba ORCAMENTOS: A-C preenchidas pelo sistema; D-F digitadas pela equipe (nunca são gravadas aqui).
 *  Cria a aba se não existir, acrescenta no fim as obras novas e atualiza endereço e cliente das existentes. Não apaga linhas. */
function sincronizarOrcamentos_(ss, linhas) {
  let sh = ss.getSheetByName(CTRL.orcamentos);
  if (!sh) {
    sh = ss.insertSheet(CTRL.orcamentos, 2);
    sh.getRange(1, 1, 1, 6).setValues([['ID_OBRA', 'OBRA (ENDEREÇO)', 'CLIENTE', 'ORÇAMENTO DE MATERIAIS (R$)', 'DATA DO ORÇAMENTO', 'OBSERVAÇÕES']])
      .setFontWeight('bold').setBackground('#EDF6DD').setFontColor('#064B25');
    sh.setFrozenRows(1);
    sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setNumberFormat('"R$ "#,##0.00');
    sh.getRange(2, 5, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 340); sh.setColumnWidth(3, 240); sh.setColumnWidth(4, 200); sh.setColumnWidth(5, 150); sh.setColumnWidth(6, 300);
  }
  const porId = {};
  linhas.forEach(l => { const id = texto_(l[0]).toUpperCase(); if (id) porId[id] = l; });
  const n = Math.max(0, sh.getLastRow() - 1), vistos = {};
  if (n) {
    const atuais = sh.getRange(2, 1, n, 3).getValues();
    let mudou = false;
    const novas = atuais.map(r => {
      const id = texto_(r[0]).toUpperCase(), l = porId[id]; if (id) vistos[id] = true;
      if (!l || (texto_(r[1]) === texto_(l[2]) && texto_(r[2]) === texto_(l[1]))) return r;
      mudou = true; return [r[0], texto_(l[2]), texto_(l[1])];
    });
    if (mudou) sh.getRange(2, 1, n, 3).setValues(novas); // só as colunas A-C
  }
  const faltam = Object.keys(porId).filter(id => !vistos[id]).map(id => [id, texto_(porId[id][2]), texto_(porId[id][1])]);
  if (!faltam.length) return 0;
  const ini = n + 2, fim = ini + faltam.length - 1;
  if (sh.getMaxRows() < fim) sh.insertRowsAfter(sh.getMaxRows(), fim - sh.getMaxRows());
  sh.getRange(ini, 1, faltam.length, 1).setNumberFormat('@');
  sh.getRange(ini, 1, faltam.length, 3).setValues(faltam).setBackground('#EEF2EF');
  sh.getRange(ini, 4, faltam.length, 3).setBackground('#FFFFFF');
  sh.getRange(ini, 4, faltam.length, 1).setNumberFormat('"R$ "#,##0.00');
  sh.getRange(ini, 5, faltam.length, 1).setNumberFormat('dd/mm/yyyy');
  return faltam.length;
}

/** Chamado pelo painel: grava o orçamento de uma obra (valor vazio = remover). Só toca D-H da linha da obra.
 *  Responde {ok:true, orcamento:{valor,data,obs}} ou {ok:false, motivo}. */
function salvarOrcamento(token, obra, valor, data, obs) {
  const ac = autenticar_(token);
  if (ac.erro) return {ok: false, acesso: false, motivo: ac.erro.motivo === 'DESATIVADO' ? 'Seu acesso ao painel está desativado.' : 'Sua sessão neste aparelho terminou. Entre de novo.'};
  if (!ehAdmin_(ac.u)) return {ok: false, motivo: 'Somente administradores podem definir ou editar orçamentos.'};
  const id = texto_(obra).toUpperCase();
  let num = '';
  if (texto_(valor) !== '') {
    num = valor_(valor);
    if (num === '' || !(num > 0) || num > 100000000) return {ok: false, motivo: 'Informe um valor maior que zero e de até R$ 100.000.000,00, ou apague o valor para remover o orçamento.'};
    num = Math.round(num * 100) / 100;
  }
  const dt = texto_(data) || Utilities.formatDate(new Date(), CTRL.fuso, 'dd/MM/yyyy');
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(dt);
  const dataCel = m ? new Date(+m[3], +m[2] - 1, +m[1], 12) : null;
  if (!m || dataCel.getMonth() !== +m[2] - 1 || dataCel.getDate() !== +m[1]) return {ok: false, motivo: 'Informe uma data válida no formato dd/mm/aaaa.'};
  const nota = texto_(obs);
  if (nota.length > 300) return {ok: false, motivo: 'A observação pode ter no máximo 300 caracteres.'};
  if (/^[=+@]/.test(nota)) return {ok: false, motivo: 'A observação não pode começar com =, + ou @.'};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return {ok: false, motivo: 'O sistema está sincronizando. Tente de novo em alguns segundos.'};
  try {
    const ss = planilhaControle_();
    const achar = () => { const sh = ss.getSheetByName(CTRL.orcamentos), n = sh ? sh.getLastRow() - 1 : 0;
      return n > 0 ? sh.getRange(2, 1, n, 1).getValues().findIndex(r => texto_(r[0]).toUpperCase() === id) : -1; };
    let i = achar();
    // Aba ou obra ainda não criadas pela sincronização: acrescenta agora (só A-C).
    if (i < 0) { sincronizarOrcamentos_(ss, linhas_(ss, CTRL.obras, CTRL.linhaObras, 3).filter(r => texto_(r[0]))); i = achar(); }
    if (i < 0) return {ok: false, motivo: 'A obra ' + id + ' não foi encontrada na aba ORCAMENTOS.'};
    const sh = ss.getSheetByName(CTRL.orcamentos), linha = i + 2;
    if (texto_(sh.getRange(1, 7).getValue()) !== 'ALTERADO POR') sh.getRange(1, 7, 1, 2).setValues([['ALTERADO POR', 'ALTERADO EM']])
      .setFontWeight('bold').setBackground('#EDF6DD').setFontColor('#064B25');
    sh.getRange(linha, 6, 1, 3).setNumberFormat('@');
    sh.getRange(linha, 4, 1, 5).setValues([[num, dataCel, nota, ac.u.nome || ac.u.usuario, Utilities.formatDate(new Date(), CTRL.fuso, 'dd/MM/yyyy HH:mm')]]);
    SpreadsheetApp.flush();
    return {ok: true, orcamento: {valor: num === '' ? null : num, data: dt, obs: nota}};
  } catch (e) { return {ok: false, motivo: 'Não foi possível salvar: ' + e.message}; }
  finally { lock.releaseLock(); }
}

/** [Função 5] Resumo semanal por e-mail: toda segunda, por volta das 7h (gatilho criado por instalarSincronizacao).
 *  Destinatários em ScriptProperties "RESUMO_DESTINOS" (gerenciados no painel, só Admin Zayda). Lista vazia = não envia nada.
 *  O conteúdo sai dos mesmos dados do painel (dadosPainel_), com as mesmas regras de custo, orçamento, parcelas e fechamento.
 *  Permissão de e-mail: rodar autorizarEmail() uma vez no editor. */
const RESUMO = Object.freeze({destinos: 'RESUMO_DESTINOS', ultimo: 'RESUMO_ULTIMO', hora: 7, max: 30});
const emailValido_ = e => /^[^\s@,;<>"']+@[^\s@,;<>"']+\.[^\s@,;<>"']{2,}$/.test(String(e || ''));
function destinosResumo_() {
  try { const l = JSON.parse(PropertiesService.getScriptProperties().getProperty(RESUMO.destinos) || '[]'); return Array.isArray(l) ? l.filter(emailValido_) : []; } catch (e) { return []; }
}
function ultimoResumo_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(RESUMO.ultimo) || 'null'); } catch (e) { return null; } }
/** Permissão de enviar e-mail (rodar uma vez no editor; termina sem erro, sem tela). */
function autorizarEmail() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, ['https://www.googleapis.com/auth/script.send_mail']);
  console.log('Permissão de e-mail autorizada: o resumo semanal já pode ser enviado.');
  return 'ok';
}
function salvarDestinosResumo(token, lista) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const brutos = (Array.isArray(lista) ? lista : []).map(e => texto_(e).toLowerCase()).filter(Boolean);
  const ruins = brutos.filter(e => !emailValido_(e));
  if (ruins.length) return {ok: false, motivo: 'E-mail inválido: ' + ruins.slice(0, 3).join(', ') + '.'};
  const l = uniq_(brutos);
  if (l.length > RESUMO.max) return {ok: false, motivo: 'No máximo ' + RESUMO.max + ' destinatários.'};
  PropertiesService.getScriptProperties().setProperty(RESUMO.destinos, JSON.stringify(l));
  return {ok: true, destinos: l};
}
function enviarResumoTeste(token) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const r = enviarResumo_('teste');
  return Object.assign(r, {ultimo: ultimoResumo_()});
}
/** Gatilho semanal (segunda, por volta das 7h). */
function resumoSemanal() { return enviarResumo_('semanal'); }
function enviarResumo_(origem) {
  const destinos = destinosResumo_(), props = PropertiesService.getScriptProperties();
  if (!destinos.length) return {ok: false, motivo: 'Nenhum destinatário cadastrado. Nada foi enviado.'};
  const em = Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm');
  try {
    try { sincronizarSeHouverNovidade_(); } catch (e) {}
    const r = montarResumo_(dadosPainel_(null));
    MailApp.sendEmail({to: destinos.join(','), subject: r.assunto, htmlBody: htmlResumo_(r), name: 'Zayda Construtora'});
    props.setProperty(RESUMO.ultimo, JSON.stringify({ok: true, em, origem, n: destinos.length}));
    return {ok: true, em, n: destinos.length};
  } catch (e) {
    console.error('Resumo semanal não enviado: ' + (e && e.message || e)); // texto técnico só no registro de execuções
    props.setProperty(RESUMO.ultimo, JSON.stringify({ok: false, em, origem}));
    return {ok: false, motivo: 'Não foi possível enviar o resumo. Tente de novo mais tarde.'};
  }
}
const diaMais_ = (iso, n) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3] + n)).toISOString().slice(0, 10) : ''; };
const br_ = iso => iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '—';
/** Os números do resumo, com as mesmas regras do painel. */
function montarResumo_(d) {
  const hoje = d.hoje, ini = diaMais_(hoje, -6), ate7 = diaMais_(hoje, 7), r2 = v => Math.round(v * 100) / 100;
  const nomeObra = id => { const o = (d.obras || []).find(x => x.id === id); return o ? (o.endereco || o.casa || id) : (id && id !== 'SEM VÍNCULO' ? id : 'Sem obra identificada'); };
  const canc = id => { const c = ((d.pedidos || {})[id] || {}).cancelamento; return c ? c.tipo : ''; };
  const ter = d.terceiros || [];
  const lanc = (d.lanc || []).map(l => Object.assign({}, l, {obra: l.obra || 'SEM VÍNCULO'}))
    .concat(ter.filter(t => t.conferencia === 'CONFERIDO' && t.dataRef && t.valor > 0).map(t => ({data: t.dataRef, obra: t.obra || 'SEM VÍNCULO', valor: t.valor, situacao: 'CUSTO CONFIRMADO', ter: true, alertas: ''})));
  const conf = lanc.filter(l => l.situacao === 'CUSTO CONFIRMADO');
  const semana = conf.filter(l => l.data >= ini && l.data <= hoje), porObra = {};
  semana.forEach(l => { porObra[l.obra] = (porObra[l.obra] || 0) + (l.valor || 0); });
  const itens = (d.itens || []).filter(i => canc(i.pedido) !== 'CANCELADO').map(i => Object.assign({}, i, {saldo: canc(i.pedido) === 'ENCERRADO' ? 0 : i.saldo}));
  // Orçamento (mesma regra da Função 1)
  const gasto = {}, aEnt = {}; conf.forEach(l => { gasto[l.obra] = (gasto[l.obra] || 0) + (l.valor || 0); });
  itens.forEach(i => { if (i.precoRef != null && i.saldo > 0) aEnt[i.obra] = (aEnt[i.obra] || 0) + i.saldo * i.precoRef; });
  const alertas = []; (d.obras || []).forEach(o => { const od = (d.orcamentos || {})[o.id], orc = od && od.valor > 0 ? od.valor : null; if (orc == null) return;
    const g = r2(gasto[o.id] || 0), u = g / orc, niv = u >= ALERTA_ORC.estouro ? 'ESTOURADO' : u >= ALERTA_ORC.atencao ? 'ATENCAO' : '';
    if (niv) alertas.push({id: o.id, nome: nomeObra(o.id), nivel: niv, pct: Math.round(u * 100), gasto: g, orc, passou: r2(Math.max(0, g - orc)), previsto: r2(g + (aEnt[o.id] || 0))}); });
  alertas.sort((a, b) => b.pct - a.pct);
  // Parcelas (mesma regra da Função 2)
  const venc = [], prox = [];
  ter.filter(t => t.conferencia !== 'CANCELADO' && t.conferencia !== 'REPROVADO').forEach(t => (t.parcelas || []).forEach((p, i) => {
    if (p.pago || !p.vencimento) return; const x = {forn: t.fornecedor, obra: t.obra, desc: t.descricao, n: p.numero || i + 1, de: (t.parcelas || []).length, valor: p.valor, venc: p.vencimento};
    if (p.vencimento < hoje) venc.push(x); else if (p.vencimento <= ate7) prox.push(x); }));
  const ordV = (a, b) => a.venc.localeCompare(b.venc); venc.sort(ordV); prox.sort(ordV);
  // Meses fechados que mudaram (mesma regra da Função 4)
  const porMes = {}; conf.forEach(l => { const k = String(l.data || '').slice(0, 7); if (!k) return; const m = porMes[k] = porMes[k] || {total: 0, obras: {}}; m.total += l.valor || 0; m.obras[l.obra] = (m.obras[l.obra] || 0) + (l.valor || 0); });
  const ultimoF = {}; (d.fechamentos || []).forEach(f => { ultimoF[f.mes] = f; });
  const mudaram = Object.keys(ultimoF).filter(k => ultimoF[k].tipo === 'FECHAMENTO').map(k => { const f = ultimoF[k], a = porMes[k] || {total: 0, obras: {}}, obras = [];
    uniq_(Object.keys(f.porObra || {}).concat(Object.keys(a.obras))).forEach(o => { if (Math.abs(((f.porObra || {})[o] || 0) - (a.obras[o] || 0)) >= 0.005) obras.push(o); });
    return Math.abs(f.total - a.total) >= 0.005 || obras.length ? {mes: k, fechado: f.total, atual: r2(a.total), obras} : null; }).filter(Boolean);
  // Demais pendências (em todo o período)
  const pedSaldo = uniq_(itens.filter(i => i.saldo > 0).map(i => i.pedido)).length;
  const pend = [['Materiais que ainda não chegaram (pedidos)', pedSaldo], ['Pedidos sem obra reconhecida', uniq_(itens.filter(i => !i.obra || i.obra === 'SEM VÍNCULO').map(i => i.pedido)).length],
    ['Entregas sem valor', lanc.filter(l => l.situacao === 'PENDENTE DE CUSTO').length], ['Entregas para conferir', lanc.filter(l => l.situacao === 'AGUARDANDO CONFERÊNCIA').length],
    ['Entregas com aviso', lanc.filter(l => l.alertas && l.situacao !== 'CANCELADO').length], ['Materiais pedidos sem preço (itens)', itens.filter(i => i.precoRef == null).length],
    ['Compras de outros fornecedores a conferir', ter.filter(t => t.conferencia === 'AGUARDANDO').length]].filter(x => x[1] > 0);
  return {assunto: 'Zayda · Resumo semanal — ' + br_(ini).slice(0, 5) + ' a ' + br_(hoje).slice(0, 5), ini, hoje,
    total: r2(sum_(semana)), obras: Object.keys(porObra).map(id => ({id, nome: nomeObra(id), v: r2(porObra[id])})).sort((a, b) => b.v - a.v),
    alertas, venc, prox, totVenc: r2(venc.reduce((s, x) => s + x.valor, 0)), totProx: r2(prox.reduce((s, x) => s + x.valor, 0)), mudaram, pend,
    link: (() => { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } })()};
}
const sum_ = a => a.reduce((s, l) => s + (l.valor || 0), 0);
const escH_ = s => String(s == null ? '' : s).replace(/[&<>"']/g, ch => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'})[ch]);
const brl_ = v => 'R$ ' + (Math.round((Number(v) || 0) * 100) / 100).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
/** E-mail em HTML simples, fundo claro, legível no celular. */
function htmlResumo_(r) {
  const V = '#00632B', T = '#13241A', C = '#4D5E54', L = '#E0E7E2';
  const sec = (tit, corpo) => '<tr><td style="padding:18px 20px 4px"><div style="font:800 15px Arial,sans-serif;color:' + T + ';margin:0 0 8px">' + tit + '</div>' + corpo + '</td></tr>';
  const lin = (a, b, cor) => '<tr><td style="padding:7px 0;border-bottom:1px solid ' + L + ';font:14px Arial,sans-serif;color:' + T + '">' + a + '</td><td style="padding:7px 0 7px 10px;border-bottom:1px solid ' + L + ';font:700 14px Arial,sans-serif;color:' + (cor || T) + ';text-align:right;white-space:nowrap">' + b + '</td></tr>';
  const tab = linhas => '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse">' + linhas.join('') + '</table>';
  const nada = t => '<div style="font:14px Arial,sans-serif;color:' + C + '">' + t + '</div>';
  const parc = x => lin(escH_(x.forn) + '<br><span style="font-size:12px;color:' + C + '">' + escH_(x.obra || 'Sem obra') + ' · parcela ' + x.n + ' de ' + x.de + ' · vence ' + br_(x.venc) + '</span>', brl_(x.valor));
  let h = '<!doctype html><html><body style="margin:0;background:#F4F7F5"><table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#F4F7F5"><tr><td align="center" style="padding:16px 8px">'
    + '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:600px;background:#FFFFFF;border:1px solid ' + L + ';border-radius:14px">'
    + '<tr><td style="padding:18px 20px;background:' + V + ';border-radius:14px 14px 0 0"><div style="font:800 18px Arial,sans-serif;color:#FFFFFF">ZAYDA CONSTRUTORA</div><div style="font:14px Arial,sans-serif;color:#E3F0E6;margin-top:4px">Resumo semanal · ' + br_(r.ini) + ' a ' + br_(r.hoje) + '</div></td></tr>';
  h += sec('Custo confirmado nos últimos 7 dias', r.total > 0 ? '<div style="font:800 24px Arial,sans-serif;color:' + V + ';margin-bottom:6px">' + brl_(r.total) + '</div>' + tab(r.obras.map(o => lin(escH_(o.id === 'SEM VÍNCULO' ? 'Sem obra identificada' : o.id + ' · ' + o.nome), brl_(o.v)))) : nada('Sem movimento nesta semana.'));
  h += sec('Orçamento das obras', r.alertas.length ? tab(r.alertas.map(a => lin((a.nivel === 'ESTOURADO' ? '▲ Estourado' : '! Atenção') + ' · ' + escH_(a.id + ' · ' + a.nome) + '<br><span style="font-size:12px;color:' + C + '">' + brl_(a.gasto) + ' de ' + brl_(a.orc) + (a.passou > 0 ? ' · passou ' + brl_(a.passou) : '') + (a.previsto > a.orc && a.previsto > a.gasto ? ' · previsto com o que falta entregar: ' + brl_(a.previsto) : '') + '</span>', a.pct + '%', a.nivel === 'ESTOURADO' ? '#B42318' : '#946000'))) : nada('Nenhuma obra em atenção ou com orçamento estourado ✓'));
  h += sec('Parcelas de outros fornecedores', (r.venc.length ? '<div style="font:700 13px Arial,sans-serif;color:#B42318;margin:4px 0">Vencidas · ' + r.venc.length + ' · ' + brl_(r.totVenc) + '</div>' + tab(r.venc.slice(0, 15).map(parc)) : '')
    + (r.prox.length ? '<div style="font:700 13px Arial,sans-serif;color:#946000;margin:10px 0 4px">Vencem nos próximos 7 dias · ' + r.prox.length + ' · ' + brl_(r.totProx) + '</div>' + tab(r.prox.slice(0, 15).map(parc)) : '')
    + (!r.venc.length && !r.prox.length ? nada('Nenhum pagamento vencendo ✓') : ''));
  if (r.mudaram.length) h += sec('Meses fechados que mudaram', tab(r.mudaram.map(m => lin(escH_(m.mes.slice(5, 7) + '/' + m.mes.slice(0, 4)) + '<br><span style="font-size:12px;color:' + C + '">Fechado ' + brl_(m.fechado) + ' · agora ' + brl_(m.atual) + (m.obras.length ? ' · obras: ' + escH_(m.obras.join(', ')) : '') + '</span>', brl_(m.atual - m.fechado), '#946000'))));
  h += sec('Outras pendências', r.pend.length ? tab(r.pend.map(p => lin(escH_(p[0]), String(p[1])))) : nada('Tudo em dia ✓'));
  if (r.link) h += '<tr><td align="center" style="padding:20px"><a href="' + escH_(r.link) + '" style="display:inline-block;background:' + V + ';color:#FFFFFF;text-decoration:none;font:800 15px Arial,sans-serif;padding:12px 22px;border-radius:10px">Abrir painel</a></td></tr>';
  h += '<tr><td style="padding:0 20px 18px;font:12px Arial,sans-serif;color:' + C + '">Enviado automaticamente pelo painel de custos da Zayda. Para sair da lista, peça a um administrador da Zayda.</td></tr></table></td></tr></table></body></html>';
  return h;
}

/** [Função 4] Fechamento do mês: fotografia dos números de um mês encerrado (aba FECHAMENTOS da planilha do painel) + PDF no Drive.
 *  O painel não altera a planilha do formulário: a "trava" é a conferência — a cada carga o navegador compara o mês atual com a fotografia.
 *  Uma linha por evento (FECHAMENTO ou REABERTURA), nada é apagado: o histórico fica todo na aba. A aba entra no backup semanal
 *  (o backup copia a planilha do painel inteira). */
const FECH = Object.freeze({aba: 'FECHAMENTOS', pasta: 'ZAYDA - Fechamentos',
  cab: ['MES', 'TIPO', 'TOTAL', 'LOJA_BSJ', 'OUTROS_FORNECEDORES', 'N_ENTREGAS', 'N_LANCAMENTOS', 'POR_OBRA', 'POR_ETAPA', 'POR', 'EM', 'PDF_URL', 'MOTIVO']});
function abaFechamentos_(criar) {
  const ss = planilhaControle_();
  let sh = ss.getSheetByName(FECH.aba);
  if (!sh && criar) {
    sh = ss.insertSheet(FECH.aba);
    sh.getRange(1, 1, 1, FECH.cab.length).setValues([FECH.cab]).setFontWeight('bold').setBackground('#EDF6DD').setFontColor('#064B25');
    sh.setFrozenRows(1);
  }
  return sh;
}
function lerFechamentos_() {
  const sh = abaFechamentos_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const js = v => { try { return JSON.parse(texto_(v) || '{}') || {}; } catch (e) { return {}; } };
  return sh.getRange(2, 1, sh.getLastRow() - 1, FECH.cab.length).getValues().filter(r => /^\d{4}-\d{2}$/.test(texto_(r[0]))).map(r => ({
    mes: texto_(r[0]), tipo: texto_(r[1]).toUpperCase(), total: num_(r[2]) || 0, bsj: num_(r[3]) || 0, ter: num_(r[4]) || 0,
    nEnt: num_(r[5]) || 0, nLanc: num_(r[6]) || 0, porObra: js(r[7]), porEtapa: js(r[8]), por: texto_(r[9]), em: texto_(r[10]), pdf: urlDrive_(r[11]), motivo: texto_(r[12])}));
}
function pastaFechamentos_() {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('FECH_PASTA_ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const it = DriveApp.getFoldersByName(FECH.pasta), f = it.hasNext() ? it.next() : DriveApp.createFolder(FECH.pasta);
  props.setProperty('FECH_PASTA_ID', f.getId());
  return f;
}
const estadoFech_ = (lista, mes) => { const l = lista.filter(f => f.mes === mes); return l.length ? l[l.length - 1] : null; };
/** Fecha um mês já terminado (só Admin Zayda). foto = números calculados no painel com as regras de custo atuais; pdf = relatório do mês (base64). */
function fecharMes(token, mes, foto, pdf) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const m = texto_(mes), atual = Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM');
  if (!/^\d{4}-\d{2}$/.test(m)) return {ok: false, motivo: 'Mês inválido.'};
  if (m >= atual) return {ok: false, motivo: 'Só é possível fechar um mês que já terminou.'};
  if (!foto || typeof foto !== 'object') return {ok: false, motivo: 'Dados do mês inválidos.'};
  const n = v => { const x = Number(v); return isFinite(x) ? Math.round(x * 100) / 100 : 0; };
  const mapa = o => { const r = {}; Object.keys(o && typeof o === 'object' ? o : {}).slice(0, 500).forEach(k => { r[String(k).slice(0, 80)] = n(o[k]); }); return r; };
  if (!pdf || String(pdf).length > 20e6) return {ok: false, motivo: 'O relatório em PDF não chegou. Tente de novo.'};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return {ok: false, motivo: 'O sistema está ocupado. Tente de novo em alguns segundos.'};
  try {
    const est = estadoFech_(lerFechamentos_(), m);
    if (est && est.tipo === 'FECHAMENTO') return {ok: false, motivo: 'Este mês já está fechado.', fechamentos: lerFechamentos_()};
    let url = '';
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(String(pdf)), 'application/pdf', 'Fechamento ' + m + ' — Zayda.pdf');
      url = pastaFechamentos_().createFile(blob).getUrl();
    } catch (e) { console.error('PDF do fechamento não salvo: ' + e.message); return {ok: false, motivo: 'Não foi possível salvar o PDF no Drive. Tente de novo.'}; }
    const sh = abaFechamentos_(true), linha = sh.getLastRow() + 1;
    sh.getRange(linha, 1, 1, 2).setNumberFormat('@'); sh.getRange(linha, 8, 1, 6).setNumberFormat('@');
    sh.getRange(linha, 1, 1, FECH.cab.length).setValues([[m, 'FECHAMENTO', n(foto.total), n(foto.bsj), n(foto.ter), Math.max(0, Math.round(Number(foto.nEnt) || 0)),
      Math.max(0, Math.round(Number(foto.nLanc) || 0)), JSON.stringify(mapa(foto.porObra)), JSON.stringify(mapa(foto.porEtapa)), a.u.nome || a.u.usuario,
      Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm'), url, '']]);
    SpreadsheetApp.flush();
    return {ok: true, fechamentos: lerFechamentos_()};
  } finally { lock.releaseLock(); }
}
/** Reabre um mês fechado (só Admin Zayda, com motivo). Fica registrado quem, quando e por quê. */
function reabrirMes(token, mes, motivo) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const m = texto_(mes), mot = texto_(motivo);
  if (mot.length < 5) return {ok: false, motivo: 'Informe o motivo (mínimo 5 caracteres).'};
  if (mot.length > 300 || /^[=+@-]/.test(mot)) return {ok: false, motivo: 'Motivo inválido (até 300 caracteres, sem começar com =, +, - ou @).'};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return {ok: false, motivo: 'O sistema está ocupado. Tente de novo em alguns segundos.'};
  try {
    const est = estadoFech_(lerFechamentos_(), m);
    if (!est || est.tipo !== 'FECHAMENTO') return {ok: false, motivo: 'Este mês não está fechado.', fechamentos: lerFechamentos_()};
    const sh = abaFechamentos_(true), linha = sh.getLastRow() + 1;
    sh.getRange(linha, 1, 1, FECH.cab.length).setNumberFormat('@');
    sh.getRange(linha, 1, 1, FECH.cab.length).setValues([[m, 'REABERTURA', '', '', '', '', '', '', '', a.u.nome || a.u.usuario, Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm'), '', mot]]);
    SpreadsheetApp.flush();
    return {ok: true, fechamentos: lerFechamentos_()};
  } finally { lock.releaseLock(); }
}

/** [Função 3] Área construída (m²) de cada obra. A planilha de obras não tem esse dado (e lá o painel só lê):
 *  fica guardada no próprio painel, em ScriptProperties "AREAS_OBRAS" = {OB-xx: {m2, por, em}}. Mesmo perfil que define o orçamento. */
const CHAVE_AREAS = 'AREAS_OBRAS';
function areasObras_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(CHAVE_AREAS) || '{}') || {}; } catch (e) { return {}; } }
function salvarArea(token, obra, m2) {
  const ac = autenticar_(token);
  if (ac.erro) return {ok: false, acesso: false, motivo: ac.erro.motivo === 'DESATIVADO' ? 'Seu acesso ao painel está desativado.' : 'Sua sessão neste aparelho terminou. Entre de novo.'};
  if (!ehAdmin_(ac.u)) return {ok: false, motivo: 'Somente administradores podem informar a área da obra.'};
  const id = texto_(obra).toUpperCase();
  if (!/^OB-\d+$/.test(id)) return {ok: false, motivo: 'Obra inválida.'};
  let v = '';
  if (texto_(m2) !== '') {
    v = valor_(m2);
    if (v === '' || !(v > 0) || v > 1000000) return {ok: false, motivo: 'Informe a área em m², maior que zero (ex.: 120,5), ou apague o valor para remover.'};
    v = Math.round(v * 100) / 100;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return {ok: false, motivo: 'O sistema está ocupado. Tente de novo em alguns segundos.'};
  try {
    const a = areasObras_();
    if (v === '') delete a[id];
    else a[id] = {m2: v, por: ac.u.nome || ac.u.usuario, em: Utilities.formatDate(new Date(), CTRL.fuso, 'dd/MM/yyyy HH:mm')};
    PropertiesService.getScriptProperties().setProperty(CHAVE_AREAS, JSON.stringify(a));
    return {ok: true, area: a[id] || null};
  } finally { lock.releaseLock(); }
}

/** Menu Zayda: cria a aba agora (sem esperar a próxima sincronização) e acrescenta as obras que faltam. */
function criarAbaOrcamentos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Outra sincronização está em andamento. Tente de novo em instantes.');
  try {
    const ss = planilhaControle_();
    const linhas = linhas_(ss, CTRL.obras, CTRL.linhaObras, 3).filter(r => texto_(r[0]));
    const novas = sincronizarOrcamentos_(ss, linhas);
    ss.toast('Aba ORCAMENTOS pronta. Obras acrescentadas: ' + novas + '.', 'Orçamentos');
    return novas;
  } finally { lock.releaseLock(); }
}

/* ===== Identificação da obra a partir do que foi digitado no formulário ===== */
function semAcento_(t) { return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); }
function palavras_(t) { return ' ' + semAcento_(t).replace(/[^a-z0-9]+/g, ' ').trim() + ' '; }
function numCasa_(t) { // "Casa 01" -> "1", "Casa 3 B" -> "3b", "Apart. 2" -> "2"
  // Letra junto do número só conta se não for início de sigla ("Casa 3 B" = 3b; "Casa 2, P.D.I" = 2).
  const m = / (?:casa|cs|apart|apto|ap|apartamento|suite) ?0*(\d+)(?: ?([a-z])(?! [a-z] ))?(?= )/.exec(palavras_(t));
  return m ? m[1] + (m[2] && m[2].length === 1 && !/^(e|a|o)$/.test(m[2]) ? m[2] : '') : '';
}
function numLote_(t) { const m = / lote ?0*(\d+[a-z]?)(?= )/.exec(palavras_(t)); return m ? m[1] : ''; }
function numPorta_(t) { const m = / (?:n|no|nº|num|numero) ?0*(\d+)(?= )/.exec(palavras_(String(t).replace(/nº/gi, ' numero ')));
  return m ? m[1] : ''; }
function rua_(endereco) { // trecho antes da 1ª vírgula ou " - ", sem "Rua/Avenida"
  const p = semAcento_(String(endereco || '').split(/,| - /)[0]).replace(/[^a-z0-9]+/g, ' ').trim();
  if (!/^(rua|r|avenida|av|travessa|tv|estrada|alameda|rodovia)\b/.test(p)) return '';
  return p.replace(/^(rua|r|avenida|av|travessa|tv|estrada|alameda|rodovia)\b\s*/, '').trim();
}
/** Endereço comparável: sem acento, minúsculo, só palavras e números, sem zero à esquerda ("Casa 02" = "casa 2"). */
function normEnd_(t) { return palavras_(t).replace(/ 0+(\d)/g, ' $1'); }
/** obras: [[id, endereco, cliente], ...]; devolve índice para resolverObra_. */
function indiceObras_(obras) {
  const porId = {}, lista = [];
  obras.forEach(o => {
    const id = String(o[0]).toUpperCase(); porId[id.replace(/[^A-Z0-9]/g, '').replace(/^OB0*/, 'OB')] = id;
    const r = rua_(o[1]);
    const full = normEnd_(o[1]);
    // Palavras do cliente que não estão no endereço: desempatam obras diferentes no mesmo endereço.
    const cli = palavras_(o[2] || '').trim().split(' ').filter(w => w.length >= 3 && !/^\d+$/.test(w) && full.indexOf(' ' + w + ' ') < 0);
    lista.push({id, rua: r ? ' ' + r + ' ' : '', casa: numCasa_(o[1]), lote: numLote_(o[1]), porta: numPorta_(o[1]),
                full: full.trim().split(' ').length >= 3 ? full : '', cli});
  });
  return {porId, lista};
}
/** Obras diferentes com o mesmo endereço (ex.: OB-76 e OB-80): fica a única cujo cliente aparece no texto. */
function desempateCliente_(cand, t) {
  if (cand.length < 2) return cand;
  const ok = cand.filter(o => o.cli.some(w => t.indexOf(' ' + w + ' ') >= 0 && !cand.some(x => x !== o && x.cli.indexOf(w) >= 0)));
  return ok.length === 1 ? ok : cand;
}
/** Código OB-xx em qualquer campo; senão endereço completo, rua + casa (+ lote); cliente desempata.
 *  '' quando não há ou é ambíguo. */
function resolverObra_(casa, endereco, idx) {
  const texto = String(casa || '') + ' ' + String(endereco || '');
  const cod = /\bOB\s*[-–.]?\s*0*(\d{1,4})\b/i.exec(texto);
  if (cod && idx.porId['OB' + cod[1]]) return idx.porId['OB' + cod[1]];
  const c = numCasa_(casa) || numCasa_(endereco) || (/^\s*0*(\d+)\s*$/.exec(String(casa || '')) || [])[1] || '';
  if (numCasa_(casa) && numCasa_(endereco) && numCasa_(casa) !== numCasa_(endereco)) return ''; // casa diferente nos dois campos
  // 1º: endereço completo da obra escrito no formulário (ex.: "Rua Frederico Silva Souto, Lote 118
  // (atrás da subprefeitura), Casa 2"). Vale o endereço mais longo encontrado, se for de uma obra só.
  const textos = [normEnd_(String(endereco || '') + ' ' + String(casa || '')), normEnd_(texto)];
  let ex = idx.lista.filter(o => o.full && textos.some(x => x.indexOf(o.full) >= 0));
  if (ex.length) {
    const maior = Math.max.apply(null, ex.map(o => o.full.length));
    ex = desempateCliente_(ex.filter(o => o.full.length === maior), palavras_(texto));
    if (ex.length === 1 && (!c || ex[0].casa === c)) return ex[0].id;
  }
  const t = palavras_(texto);
  const l = numLote_(texto), n = numPorta_(texto);
  let cand = idx.lista.filter(o => o.rua && t.indexOf(o.rua) >= 0);
  if (!cand.length && l) cand = idx.lista.filter(o => !o.rua && o.lote === l); // endereços sem rua: pelo lote
  if (!cand.length) return '';
  if (c) cand = cand.filter(o => o.casa === c);
  if (l) { const mesmo = cand.filter(o => o.lote === l); cand = mesmo.length ? mesmo : cand.filter(o => !o.lote); }
  let porta = n;
  if (!porta && cand.length && cand[0].rua) { const m = new RegExp(cand[0].rua + '0*(\\d+) ').exec(t); if (m) porta = m[1]; } // "Rua São João 662"
  if (porta) { const mesma = cand.filter(o => o.porta === porta); if (mesma.length) cand = mesma; }
  if (cand.length > 1 && !l) { const semLote = cand.filter(o => !o.lote); if (semLote.length === 1) cand = semLote; }
  cand = desempateCliente_(cand, t);
  const ids = [...new Set(cand.map(o => o.id))];
  return ids.length === 1 ? ids[0] : '';
}

/** Mesma chave da planilha: minúsculas, só letras e números, sem zeros à esquerda ("Casa 04" -> "casa4"). */
function chave_(v) {
  return texto_(v).toLowerCase().replace(/[^a-z0-9]/g, '').replace(/(^|[a-z])0+([0-9])/g, '$1$2');
}

/** ID_OBRA de cada linha de pedido: nome exato (OBRAS ou nomes alternativos) ou, senão, código OB-xx
 *  ou rua + casa/lote digitados no endereço e no campo Casa/obra. Devolve {porLinha:[...], vigente:{pedido:id}}. */
function identificarObras_(ss, dados, obrasLinhas) {
  const mapa = {};
  (obrasLinhas || []).forEach(o => { mapa[chave_(o[3] || o[0])] = o[0]; });
  const conf = ss.getSheetByName(CTRL.conf);
  const nAp = conf.getLastRow() - CTRL.linhaApelidos + 1;
  if (nAp > 0) conf.getRange(CTRL.linhaApelidos, 1, nAp, 2).getValues()
    .forEach(r => { if (texto_(r[0]) && texto_(r[1])) mapa[chave_(r[0])] = texto_(r[1]); });
  const idx = indiceObras_((obrasLinhas || []).map(o => [o[0], o[2], o[1]]));
  const resolver = p => { try { return resolverObra_(p[8], p[7], idx); } catch (e) { return ''; } }; // um texto estranho nunca trava a sincronização
  const porLinha = dados.pedidos.map(p => mapa[chave_(p[8])] || resolver(p) || '');
  const vigente = {}, revMax = {};
  dados.pedidos.forEach((p, i) => { if (!(p[0] in revMax) || p[1] > revMax[p[0]]) { revMax[p[0]] = p[1]; vigente[p[0]] = porLinha[i]; } });
  return {porLinha, vigente, revMax};
}

/** Acrescenta em MATERIAIS_E_CUSTOS as entregas do formulário que ainda não estão lá. */
function lancarEntregas_(ss, dados, ident) {
  if (!dados.entregas.length) return 0;
  const sh = ss.getSheetByName(CTRL.lancamentos);
  const ult = sh.getLastRow();
  const colA = ult >= CTRL.linhaLanc ? sh.getRange(CTRL.linhaLanc, 1, ult - CTRL.linhaLanc + 1, 1).getValues() : [];
  const existentes = {};
  let ultimaComId = CTRL.linhaLanc - 1;
  colA.forEach((r, i) => { const v = texto_(r[0]); if (v) { existentes[v.toUpperCase()] = true; ultimaComId = CTRL.linhaLanc + i; } });
  const novas = dados.entregas.filter(e => !existentes[e.id.toUpperCase()]);
  const obraDe = pedido => ident.vigente[pedido] || '';
  const revMax = ident.revMax;
  // Preço unitário do pedido (revisão vigente) por pedido + SKU: vira o custo real da entrega.
  const preco = {};
  dados.itens.filter(i => i[1] === revMax[i[0]]).forEach(i => { preco[i[0] + '|' + i[2]] = i[7]; });
  const precoDe = e => { const v = preco[e.pedido + '|' + e.sku]; return typeof v === 'number' && isFinite(v) ? v : ''; };
  // Linhas ENT- já lançadas sem obra (pedido sem código na época): completa a obra quando o pedido
  // passa a ter uma obra reconhecida. Só preenche a coluna C quando ela está vazia.
  if (colA.length) {
    const porEnt = {}; dados.entregas.forEach(e => { porEnt[e.id.toUpperCase()] = e; });
    const colC = sh.getRange(CTRL.linhaLanc, 3, colA.length, 1).getValues();
    colA.forEach((r, i) => {
      const e = porEnt[texto_(r[0]).toUpperCase()];
      if (!e || texto_(colC[i][0])) return;
      const ob = obraDe(e.pedido);
      if (ob) sh.getRange(CTRL.linhaLanc + i, 3).setValue(ob);
    });
  }
  if (!novas.length) return 0;
  const ini = ultimaComId + 1;
  if (sh.getMaxRows() < ini + novas.length - 1) sh.insertRowsAfter(sh.getMaxRows(), ini + novas.length - 1 - sh.getMaxRows());
  // Data como texto dd/mm/aaaa: a planilha (pt_BR) converte em data sem depender do fuso do script.
  const dataCel = d => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d); return m ? m[3] + '/' + m[2] + '/' + m[1] : d; };
  // Só colunas de digitação: G, H, K, O e P são fórmulas da planilha e não podem receber valores.
  sh.getRange(ini, 1, novas.length, 6).setValues(novas.map(e => [e.id, dataCel(e.data), obraDe(e.pedido),
    'SOLICITAÇÃO', e.pedido, e.sku]));
  sh.getRange(ini, 9, novas.length, 2).setValues(novas.map(e => [e.qtd, precoDe(e)]));
  sh.getRange(ini, 12, novas.length, 3).setValues(novas.map(e => [e.email, 'CONFIRMADO',
    'Entrega confirmada no formulário · ' + (precoDe(e) === '' ? 'item sem preço no pedido: informar o custo' : 'custo = preço do pedido') +
    (e.recebido ? ' · recebido por ' + e.recebido : '') + (e.obs ? ' · ' + e.obs : '')]));
  return novas.length;
}

/** Aba OBRAS na planilha do formulário: ID · ENDEREÇO · CLIENTE · STATUS · CASA sugerida.
 *  É a única aba que o controle escreve no formulário; serve só para as sugestões de endereço.
 *  Obras diferentes no mesmo endereço recebem o nome do cliente na casa sugerida, para não haver dúvida. */
function escreverObrasFormulario_(linhas) {
  const conf = planilhaControle_().getSheetByName(CTRL.conf);
  const qtd = {};
  linhas.forEach(o => { const k = normEnd_(o[2]); qtd[k] = (qtd[k] || 0) + 1; });
  const dados = linhas.filter(o => texto_(o[2])).map(o => {
    const m = /\b(casa|apart\.?|apto\.?|su[ií]te)\s*\d+\s?[a-z]?\b/i.exec(o[2]) || /\blote\s*[\w]+/i.exec(o[2]);
    let casa = m ? m[0].replace(/\s+/g, ' ') : o[0];
    if (qtd[normEnd_(o[2])] > 1 && texto_(o[1])) casa += ' · ' + texto_(o[1]); // mesmo endereço: o cliente diferencia
    else if (qtd[normEnd_(o[2])] > 1) casa += ' (' + o[0] + ')';
    return [o[0], texto_(o[2]), texto_(o[1]), texto_(o[7]), casa];
  });
  const props = PropertiesService.getScriptProperties();
  const assinatura = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(dados)));
  const origem = SpreadsheetApp.openById(texto_(conf.getRange(CTRL.celulaOrigemId).getValue()));
  let sh = origem.getSheetByName(CTRL.obras);
  if (sh && props.getProperty('HASH_OBRAS_FORMULARIO') === assinatura) return;
  if (!sh) {
    sh = origem.insertSheet(CTRL.obras, origem.getNumSheets());
    sh.protect().setWarningOnly(true).setDescription('Preenchida automaticamente pelo Controle de Obras - não editar');
  }
  sh.clearContents();
  sh.getRange(1, 1, 1, 5).setValues([['ID_OBRA', 'ENDEREÇO', 'CLIENTE', 'STATUS', 'CASA (SUGESTÃO DO FORMULÁRIO)']])
    .setFontWeight('bold').setBackground('#00632B').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  if (dados.length) sh.getRange(2, 1, dados.length, 5).setValues(dados);
  props.setProperty('HASH_OBRAS_FORMULARIO', assinatura);
}

function uniq_(a) { return a.filter((v, i) => a.indexOf(v) === i); }

function escrever_(ss, nome, linhas, nCols) {
  const sh = ss.getSheetByName(nome);
  if (!sh) throw new Error('Aba técnica não encontrada: ' + nome);
  const ultima = sh.getMaxRows();
  if (ultima > 1) sh.getRange(2, 1, ultima - 1, nCols).clearContent();
  if (linhas.length) {
    if (sh.getMaxRows() < linhas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), linhas.length + 1 - sh.getMaxRows());
    // Texto puro para IDs e SKUs: evita conversão automática de "00123" em número.
    sh.getRange(2, 1, linhas.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, linhas.length, nCols).setValues(linhas);
  }
}

/** Grava os pedidos cancelados/encerrados na aba oculta _IMPORT_CANCELAMENTOS (cria se não existir). */
function escreverCancelamentos_(ss, linhas) {
  let sh = ss.getSheetByName(CTRL.destCancel);
  if (!sh) { sh = ss.insertSheet(CTRL.destCancel, ss.getNumSheets()); sh.hideSheet(); }
  sh.getRange(1, 1, 1, 6).setValues([['PEDIDO_ID', 'TIPO', 'DATA_HORA', 'EMAIL_LOGIN', 'MOTIVO', 'ITENS_NAO_ENTREGUES']]).setFontWeight('bold');
  const ultima = sh.getMaxRows();
  if (ultima > 1) sh.getRange(2, 1, ultima - 1, 6).clearContent();
  if (!linhas.length) return;
  if (sh.getMaxRows() < linhas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), linhas.length + 1 - sh.getMaxRows());
  sh.getRange(2, 1, linhas.length, 6).setNumberFormat('@');
  sh.getRange(2, 1, linhas.length, 6).setValues(linhas);
}

/** Grava cada entrega do formulário na aba oculta _IMPORT_ENTREGAS (cria se não existir), com quem entregou e quem recebeu.
 *  Entrega sem ENTREGUE_POR (antiga) recebe o entregador do pedido (revisão mais alta). Os lançamentos não mudam. */
function escreverEntregas_(ss, entregas, pedidos) {
  const entregador = {}, rev = {};
  pedidos.forEach(p => { if (!(p[0] in rev) || p[1] > rev[p[0]]) { rev[p[0]] = p[1]; entregador[p[0]] = p[6]; } });
  // K REGISTRADO_POR: e-mail de quem confirmou a entrega no formulário (coluna I EMAIL_LOGIN da aba ENTREGAS).
  const cab = ['ENTREGA_ID', 'PEDIDO_ID', 'SKU', 'PRODUTO', 'UNIDADE', 'QTD_ENTREGUE', 'DATA_ENTREGA', 'RECEBIDO_POR', 'ENTREGUE_POR', 'OBSERVACOES', 'REGISTRADO_POR'];
  let sh = ss.getSheetByName(CTRL.destEntregas);
  if (!sh) { sh = ss.insertSheet(CTRL.destEntregas, ss.getNumSheets()); sh.hideSheet(); }
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  const ultima = sh.getMaxRows();
  if (ultima > 1) sh.getRange(2, 1, ultima - 1, cab.length).clearContent();
  if (!entregas.length) return;
  const linhas = entregas.map(e => [e.id, e.pedido, e.sku, e.produto, e.un, e.qtd, e.data, e.recebidoTxt, e.entreguePor || entregador[e.pedido] || '', e.obsTxt, texto_(e.email).toLowerCase()]);
  if (sh.getMaxRows() < linhas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), linhas.length + 1 - sh.getMaxRows());
  sh.getRange(2, 1, linhas.length, cab.length).setNumberFormat('@');
  sh.getRange(2, 6, linhas.length, 1).setNumberFormat('0.###');
  sh.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
}

/** Execução principal (também usada pelo acionador automático). */
function sincronizarFormulario(esperaMs) {
  // Trava do script (não do documento): vale também quando a sincronização é chamada pelo painel (app da Web).
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(esperaMs || 30000)) throw new Error('Outra sincronização está em andamento.');
  try {
    const ss = planilhaControle_();
    const dados = lerOrigem_(); // toda a leitura acontece antes de qualquer escrita
    // Obras: uma falha aqui não impede a sincronização dos pedidos; fica registrada no resumo.
    let obras = {ativo: false, linhas: []}, avisoObras = '';
    try { obras = lerObras_(); } catch (e) { avisoObras = ' · OBRAS NÃO ATUALIZADAS: ' + e.message; }
    const obrasAtuais = Math.max(0, ss.getSheetByName(CTRL.obras).getLastRow() - CTRL.linhaObras + 1);
    if (obras.ativo && !obras.linhas.length && obrasAtuais > 0) {
      obras.ativo = false; avisoObras = ' · OBRAS NÃO ATUALIZADAS: a planilha de obras não retornou nenhuma obra OB-xx';
    }
    const atuais = ss.getSheetByName(CTRL.destPedidos).getLastRow() - 1;
    if (!dados.pedidos.length && atuais > 0 && !dados.catalogo.length) {
      throw new Error('A origem não retornou pedidos nem catálogo. Confira o ID da planilha em CONFIGURACOES. Nada foi alterado.');
    }
    // [Etapa 1] Só regrava a aba de destino quando o conteúdo mudou (impressão digital MD5 em PropertiesService). A cada 6 h regrava tudo,
    // por segurança. Se a comparação falhar, copia normalmente.
    const props = PropertiesService.getScriptProperties(), tudo = Date.now() - Number(props.getProperty('REESCRITA_TOTAL') || 0) > 6 * 3600000, digs = {};
    const gravar = (chave, conteudo, fn) => { let d = '';
      try { d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(conteudo), Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); } catch (e) { d = ''; }
      if (d && !tudo && props.getProperty('DIG_' + chave) === d) return false;
      fn(); if (d) digs['DIG_' + chave] = d; return true; };
    gravar('catalogo', dados.catalogo, () => escrever_(ss, CTRL.destCatalogo, dados.catalogo, 6));
    gravar('itens', dados.itens, () => escrever_(ss, CTRL.destItens, dados.itens, 11));
    if (obras.ativo && (tudo || obras.assinatura !== props.getProperty('HASH_OBRAS'))) {
      escreverObras_(ss, obras.linhas); escreverEtapas_(ss, obras.etapas);
      PropertiesService.getScriptProperties().setProperty('HASH_OBRAS', obras.assinatura);
    }
    const obrasParaVinculo = obras.ativo ? obras.linhas :
      linhas_(ss, CTRL.obras, CTRL.linhaObras, 4).filter(r => texto_(r[0]));
    const ident = identificarObras_(ss, dados, obrasParaVinculo);
    // Coluna O de _IMPORT_PEDIDOS: obra identificada pelo script (código ou rua + casa).
    // [Função 6] Coluna Q de _IMPORT_PEDIDOS: ETAPA da obra informada no pedido.
    const linhasPed = dados.pedidos.map((p, i) => p.concat([ident.porLinha[i], p.referencia || '', p.etapa || '']));
    gravar('pedidos', linhasPed, () => escrever_(ss, CTRL.destPedidos, linhasPed, 17));
    const cabP = ss.getSheetByName(CTRL.destPedidos).getRange('P1');
    if (texto_(cabP.getValue()) !== 'PONTO_REFERENCIA') cabP.setValue('PONTO_REFERENCIA');
    const cabQ = ss.getSheetByName(CTRL.destPedidos).getRange('Q1');
    if (texto_(cabQ.getValue()) !== 'ETAPA') cabQ.setValue('ETAPA');
    gravar('cancelamentos', dados.cancelamentos, () => escreverCancelamentos_(ss, dados.cancelamentos));
    gravar('entregas', [dados.entregas, dados.pedidos.map(p => [p[0], p[1], p[6]])], () => escreverEntregas_(ss, dados.entregas, dados.pedidos));
    const novasEntregas = lancarEntregas_(ss, dados, ident);
    // Lista de obras para as sugestões de endereço do formulário (só regrava quando muda).
    if (obras.ativo) { try { escreverObrasFormulario_(obras.linhas); } catch (e) { avisoObras += ' · LISTA DE OBRAS DO FORMULÁRIO NÃO ATUALIZADA: ' + e.message; } }
    const resumo = Utilities.formatDate(new Date(), CTRL.fuso, 'dd/MM/yyyy HH:mm') + ' · ' +
      uniq_(dados.pedidos.map(p => p[0])).length + ' pedidos · ' + dados.pedidos.length + ' revisões · ' +
      dados.itens.length + ' itens · ' + dados.catalogo.length + ' SKUs' +
      (obras.ativo ? ' · ' + obras.linhas.length + ' obras' : '') + avisoObras +
      (novasEntregas ? ' · ' + novasEntregas + ' entregas novas' : '') +
      (dados.dup.length ? ' · REVISÕES DUPLICADAS: ' + dados.dup.join(', ') : '') +
      (dados.orfaos.length ? ' · ITENS SEM PEDIDO: ' + dados.orfaos.join(', ') : '');
    ss.getSheetByName(CTRL.conf).getRange(CTRL.celulaUltimaSinc).setNumberFormat('@').setValue(resumo);
    SpreadsheetApp.flush();
    try { if (tudo) digs.REESCRITA_TOTAL = String(Date.now()); if (Object.keys(digs).length) props.setProperties(digs); } catch (e) {} // [Etapa 1]
    Logger.log(resumo);
    return resumo;
  } finally { lock.releaseLock(); }
}

/** Aviso na tela da planilha (menu Zayda). Pelo editor ou por gatilho não há tela: só registra no log, sem erro. */
function avisar_(titulo, mensagem) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (titulo) ui.alert(titulo, mensagem, ui.ButtonSet.OK); else ui.alert(mensagem);
  } catch (e) {
    console.log(titulo ? titulo + ': ' + mensagem : mensagem);
  }
}

function sincronizarComAviso() {
  try { avisar_('Sincronização concluída', sincronizarFormulario()); }
  catch (e) { avisar_('Sincronização não concluída', e.message); }
}

/** Acionador de 5 em 5 minutos (CTRL.intervaloMinutos). Verificação rápida: só faz a cópia completa quando o formulário
 *  registrou algo novo (pedido, revisão ou entrega) ou a cada 15 minutos (obras e catálogo).
 *  Assim fica quase em tempo real sem estourar o limite diário de execução do Google. */
/** Marca rápida do formulário: nº de linhas de PEDIDOS, ITENS_PEDIDO, ENTREGAS, AUDITORIA, CANCELAMENTOS, TERCEIROS, TERCEIROS_PARCELAS, FORNECEDORES e CATALOGO. */
function marcaOrigem_(conf) {
  const origem = SpreadsheetApp.openById(texto_(conf.getRange(CTRL.celulaOrigemId).getValue()));
  const ult = nome => { const sh = origem.getSheetByName(nome); return sh ? sh.getLastRow() : 0; };
  // Todas as abas do formulário que o painel usa: só o número de linhas de cada uma (leve, sem ler o conteúdo).
  // Edições que não criam linha (conferência, pagamento, edição de lançamento, preço do catálogo) entram na AUDITORIA, que também é contada.
  return [ult(texto_(conf.getRange(CTRL.celulaAbaPedidos).getValue())), ult(texto_(conf.getRange(CTRL.celulaAbaItens).getValue())),
          ult(CTRL.abaEntregasOrigem), ult('AUDITORIA'), ult(CTRL.abaCancelOrigem), ult(TERC.abaLanc), ult(TERC.abaParc), ult(TERC.abaForn),
          ult(texto_(conf.getRange(CTRL.celulaAbaCatalogo).getValue()))].join('|');
}

/** Chamado pelo painel antes de devolver os dados: se o formulário tem pedido, revisão ou entrega
 *  nova, sincroniza na hora. Assim o painel mostra a novidade sem esperar o acionador de 5 minutos. */
function sincronizarSeHouverNovidade_() {
  try {
    const conf = planilhaControle_().getSheetByName(CTRL.conf);
    const props = PropertiesService.getScriptProperties();
    const marca = marcaOrigem_(conf);
    if (marca === props.getProperty('MARCA_ORIGEM')) return atualizarObrasSeMudou_(false);
    sincronizarFormulario(1500); // se outra sincronização estiver rodando, não espera: o painel lê o que já existe
    props.setProperty('MARCA_ORIGEM', marca);
    props.setProperty('ULTIMA_COPIA_COMPLETA', String(Date.now()));
    return true;
  } catch (e) { Logger.log('Sincronização pelo painel não feita: ' + e.message); return false; }
}

/** [Etapa 1] Gatilho de 5 em 5 minutos: mede a duração de cada execução (cota de 90 min/dia de gatilhos da conta gratuita). */
function sincronizacaoAutomatica() {
  const t0 = Date.now(); let r = '';
  try { r = sincronizacaoAutomaticaCorpo_(); marcarErroSinc_(''); return r; }
  catch (e) { marcarErroSinc_(String(e && e.message || e)); throw e; } // a tela Sistema mostra o aviso (só para Admin Zayda)
  finally { registrarExecucaoSinc_(Date.now() - t0, /^(Sem novidades|Outra sincronização)/.test(String(r || ''))); }
}
function sincronizacaoAutomaticaCorpo_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(200)) return 'Outra sincronização em andamento.';
  lock.releaseLock();
  const ss = planilhaControle_();
  const conf = ss.getSheetByName(CTRL.conf);
  const props = PropertiesService.getScriptProperties();
  const marca = marcaOrigem_(conf);
  const ultimaCompleta = Number(props.getProperty('ULTIMA_COPIA_COMPLETA') || 0);
  const agora = Date.now();
  if (marca === props.getProperty('MARCA_ORIGEM') && agora - ultimaCompleta < CTRL.copiaCompletaMinutos * 60000) {
    // Nada novo no formulário: confere se a planilha de obras (fase, status, etapas) mudou
    // e atualiza a hora da verificação em B10 (o painel mostra essa hora).
    const mudouObras = atualizarObrasSeMudou_(false);
    const cel = conf.getRange(CTRL.celulaUltimaSinc), txt = texto_(cel.getValue()), i = txt.indexOf(' · ');
    cel.setValue(Utilities.formatDate(new Date(), CTRL.fuso, 'dd/MM/yyyy HH:mm') + (i >= 0 ? txt.slice(i) : ''));
    return mudouObras ? 'Obras atualizadas.' : 'Sem novidades no formulário.';
  }
  const r = sincronizarFormulario();
  props.setProperty('MARCA_ORIGEM', marca);
  props.setProperty('ULTIMA_COPIA_COMPLETA', String(agora));
  return r;
}

/** everyMinutes só aceita 1, 5, 10, 15 ou 30: outro número vira o permitido logo acima (ex.: 2 → 5; acima de 30 → 30). */
function intervaloPermitido_(n) {
  const ok = [1, 5, 10, 15, 30], v = Number(n);
  if (!(v > 0)) return 5;
  for (let i = 0; i < ok.length; i++) if (ok[i] >= v) return ok[i];
  return 30;
}

function instalarSincronizacao() {
  const minutos = intervaloPermitido_(CTRL.intervaloMinutos);
  // Ordem segura: primeiro cria o gatilho novo; só depois apaga os antigos da sincronização.
  // Se a criação falhar, nada é apagado e os gatilhos antigos continuam funcionando.
  const antigos = ScriptApp.getProjectTriggers();
  const novo = ScriptApp.newTrigger('sincronizacaoAutomatica').timeBased().everyMinutes(minutos).create();
  // [Função 5] Resumo semanal por e-mail: segunda-feira, por volta das 7h (criado antes de apagar o antigo).
  const novoResumo = ScriptApp.newTrigger('resumoSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(RESUMO.hora).create();
  antigos
    .filter(t => t.getUniqueId() !== novo.getUniqueId() && t.getUniqueId() !== novoResumo.getUniqueId())
    .filter(t => ['sincronizarFormulario', 'sincronizacaoAutomatica', 'aoEditarObras', 'backupSemanal', 'resumoSemanal'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  // Na hora em que alguém muda fase, status ou etapa na planilha "Controle e Planejamento de Obras".
  const idObras = texto_(planilhaControle_().getSheetByName(CTRL.conf).getRange(CTRL.celulaObrasId).getValue());
  if (idObras) ScriptApp.newTrigger('aoEditarObras').forSpreadsheet(idObras).onEdit().create();
  // [Etapa 1] Backup semanal: segunda-feira entre 3h e 4h (fuso do projeto: São Paulo).
  ScriptApp.newTrigger('backupSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(3).create();
  PropertiesService.getScriptProperties().deleteProperty('MARCA_ORIGEM');
  const r = sincronizacaoAutomatica();
  avisar_('', 'Sincronização automática ativada: verifica o formulário a cada ' + minutos + ' minutos e copia assim que houver pedido ou entrega nova. ' +
    'Mudanças de fase, status e etapas na planilha de obras entram na hora. Backup semanal: segunda, entre 3h e 4h. Resumo por e-mail: segunda, por volta das 7h.\n\nPrimeira execução: ' + r);
}

/* ===================== REMOÇÃO DOS DADOS DE TESTE ===================== */
function ehTeste_(v) { const s = texto_(v).toUpperCase(); return s.indexOf('TST') === 0 || s.indexOf('TESTE-') === 0 || s.indexOf('T-') === 0; }

function removerLinhas_(sh, primeira, col, teste) {
  const n = sh.getLastRow() - primeira + 1;
  if (n < 1) return 0;
  const vals = sh.getRange(primeira, col, n, 1).getValues();
  let removidas = 0;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (teste(vals[i][0])) { sh.deleteRow(primeira + i); removidas++; }
  }
  return removidas;
}

function removerDadosDeTeste() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('Remover dados de teste',
    'Serão apagadas as obras TST-, os lançamentos TST-, o apelido de teste e os pedidos/itens/SKUs de teste das abas de importação. ' +
    'Registros reais não são tocados. Deseja continuar?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const ss = planilhaControle_();
  const conf = ss.getSheetByName(CTRL.conf);
  const r = {
    lancamentos: removerLinhas_(ss.getSheetByName(CTRL.lancamentos), 5, 1, ehTeste_),
    obras: removerLinhas_(ss.getSheetByName(CTRL.obras), 5, 1, ehTeste_),
    apelidos: removerLinhas_(conf, 18, 2, ehTeste_),
    pedidos: removerLinhas_(ss.getSheetByName(CTRL.destPedidos), 2, 1, ehTeste_),
    itens: removerLinhas_(ss.getSheetByName(CTRL.destItens), 2, 1, ehTeste_),
    skus: removerLinhas_(ss.getSheetByName(CTRL.destCatalogo), 2, 1, ehTeste_)
  };
  conf.getRange(CTRL.celulaModo).setValue('PRODUÇÃO');
  ui.alert('Dados de teste removidos', JSON.stringify(r), ui.ButtonSet.OK);
}

/* ========================= PAINEL (APP DA WEB) ========================= */
/** Assinatura curta do Painel.html publicado (MD5, 8 caracteres). Calculada a cada leitura, sem cache,
 *  para uma Nova versão ser percebida logo depois de publicada. */
function versaoPainel_() {
  try {
    const conteudo = HtmlService.createHtmlOutputFromFile('Painel').getContent();
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, conteudo, Utilities.Charset.UTF_8)
      .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 8);
  } catch (e) { Logger.log('Assinatura da versão não calculada: ' + e.message); return ''; }
}

function doGet() {
  const url = ScriptApp.getService().getUrl() || '';
  const html = HtmlService.createHtmlOutputFromFile('Painel').getContent()
    .split('"__URL_API__"').join(JSON.stringify(url).replace(/\//g, '\\/')) // sem "//": o removedor de comentários do HtmlService pode cortar a URL ali
    .split('"__VERSAO_PAINEL__"').join(JSON.stringify(versaoPainel_()));
  return HtmlService.createHtmlOutput(html)
    .setTitle('Zayda | Painel de custos por obra')
    .setFaviconUrl('https://res.cloudinary.com/dovqcebdt/image/upload/v1793988690/favicon-zayda_jiggwc.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let corpo = {};
  try { corpo = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { corpo = {}; }
  let saida;
  const c = corpo;
  try { saida = c.acao === 'entrar' ? entrarPainel(c.usuario, c.senha !== undefined ? c.senha : c.pin)
    : c.acao === 'verificarPrimeiro' ? verificarPrimeiroAcesso(c.usuario, c.codigo)
    : c.acao === 'primeiro' ? primeiroAcesso(c.usuario, c.codigo, c.senha !== undefined ? c.senha : c.pin)
    : c.acao === 'trocarSenha' ? trocarSenha(c.token, c.senhaAtual, c.senhaNova)
    : c.acao === 'trocarPin' ? trocarPin(c.token, c.pinAtual, c.pinNovo)
    : c.acao === 'sair' ? sairAparelho(c.token)
    : c.acao === 'usuarios' ? listarUsuarios(c.token)
    : c.acao === 'salvarUsuario' ? salvarUsuario(c.token, c.dados)
    : c.acao === 'resetarPin' ? resetarPin(c.token, c.usuario)
    : c.acao === 'desconectar' ? desconectarAparelhos(c.token, c.usuario)
    : c.acao === 'entregas' ? getEntregasPedido(c.token, c.pedido)
    : c.acao === 'salvarOrcamento' ? salvarOrcamento(c.token, c.obra, c.valor, c.data, c.obs)
    : c.acao === 'salvarArea' ? salvarArea(c.token, c.obra, c.m2)
    : c.acao === 'fecharMes' ? fecharMes(c.token, c.mes, c.foto, c.pdf)
    : c.acao === 'reabrirMes' ? reabrirMes(c.token, c.mes, c.motivo)
    : c.acao === 'conferirTerceiro' ? conferirTerceiro(c.token, c.id, c.op, c.motivo, c.atualizadoEm)
    : c.acao === 'pagarParcelaTerceiro' ? pagarParcelaTerceiro(c.token, c.id, c.op, c.data, c.forma, c.motivo)
    : c.acao === 'backup' ? backupAgora(c.token)
    : c.acao === 'destinosResumo' ? salvarDestinosResumo(c.token, c.destinos)
    : c.acao === 'resumoTeste' ? enviarResumoTeste(c.token)
    : getDadosPainel(c.token); }
  catch (err) { saida = {acesso: false, motivo: 'ERRO', detalhe: String(err && err.message || err)}; }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

function mostrarLinkPainel() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('O painel ainda não foi implantado. Use Implantar > Nova implantação > App da Web.'); return; }
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px;color:#1f2a24">' +
    '<p style="margin:0 0 12px"><a href="' + url + '" target="_blank" ' +
    'style="display:inline-block;background:#00632B;color:#fff;padding:9px 16px;border-radius:6px;text-decoration:none;font-weight:bold">' +
    'Abrir o painel</a></p>' +
    '<p style="margin:0 0 6px">Endereço para enviar à diretoria:</p>' +
    '<input value="' + url + '" readonly onclick="this.select()" style="width:100%;padding:7px;box-sizing:border-box;font-size:12px">' +
    '<p style="margin:8px 0 0;color:#5f6b64">Clique no endereço para selecionar e copie com Ctrl+C.</p></div>')
    .setWidth(480).setHeight(200);
  ui.showModalDialog(html, 'Painel executivo Zayda');
}

function linhas_(ss, aba, primeiraLinha, nCols) {
  const sh = ss.getSheetByName(aba);
  if (!sh) throw new Error('Aba obrigatória não encontrada: ' + aba);
  const n = sh.getLastRow() - primeiraLinha + 1;
  return n < 1 ? [] : sh.getRange(primeiraLinha, 1, n, nCols).getValues();
}
function dataIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CTRL.fuso, 'yyyy-MM-dd');
  const s = texto_(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s); // dd/mm/aaaa
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
}
function num_(v) { return (v === '' || v === null || v === undefined || isNaN(Number(v))) ? null : Number(v); }

/* ===================== CÓDIGO DE ACESSO DO PAINEL ===================== */
// O painel é publicado "Executar como: Eu" + "Qualquer pessoa": abre em qualquer navegador, com
// qualquer conta Google (ou nenhuma). Os números só são enviados a quem informar o código da equipe.
const CHAVE_CODIGO_PAINEL = 'CODIGO_PAINEL';

function definirCodigoPainel() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  const existe = !!props.getProperty(CHAVE_CODIGO_PAINEL);
  const r = ui.prompt('Código de acesso do painel',
    'Digite o código que a equipe vai usar para abrir o painel (mínimo 6 caracteres, letras e números).' +
    (existe ? '\n\nJá existe um código: o novo substitui o anterior e todos precisarão digitar de novo.' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const c = r.getResponseText().trim();
  if (c.length < 6) { ui.alert('Use pelo menos 6 caracteres.'); return; }
  props.setProperty(CHAVE_CODIGO_PAINEL, c);
  CacheService.getScriptCache().remove('FALHAS_PAINEL');
  ui.alert('Código definido', 'Para abrir o painel, a equipe vai digitar:\n\n' + c +
    '\n\nPasse o código só para quem pode ver os custos.', ui.ButtonSet.OK);
}

/** 'OK', 'SEM_CODIGO' (nenhum definido), 'INVALIDO' ou 'BLOQUEADO' (muitas tentativas erradas). */
function validarCodigoPainel_(codigo) {
  const certo = PropertiesService.getScriptProperties().getProperty(CHAVE_CODIGO_PAINEL);
  if (!certo) return 'SEM_CODIGO';
  const cache = CacheService.getScriptCache(), falhas = Number(cache.get('FALHAS_PAINEL') || 0);
  if (falhas >= 20) return 'BLOQUEADO';
  if (texto_(codigo).toLowerCase() === certo.toLowerCase()) return 'OK';
  cache.put('FALHAS_PAINEL', String(falhas + 1), 600);
  Utilities.sleep(1200);
  return 'INVALIDO';
}

/* ===================== ACESSO INDIVIDUAL: USUÁRIO + SENHA E CHAVE DO APARELHO =====================
 * Cadastro na aba USUARIOS da planilha do formulário: A EMAIL · B NOME · C USUARIO · D EMPRESA · E PERFIL ·
 * F FORMULARIO · G PAINEL · H ATIVO · I OBSERVACAO. A senha nunca vai para a planilha nem para o navegador:
 * fica nas propriedades do script ("PIN_<usuario>" = "v2:repetições:sal:hash"; PIN antigo: "sal:hash"). Cada aparelho conectado tem uma chave
 * ("TOKEN_<chave>" = usuário, criação e último uso). O código mestre (CODIGO_PAINEL) serve para o primeiro
 * acesso e, sozinho, para a recuperação do dono. */
const ABA_USUARIOS = 'USUARIOS';
const DONO_SISTEMA = 'vendas.zaydaimoveis1@gmail.com';
const CAB_USUARIOS = ['EMAIL', 'NOME', 'USUARIO', 'EMPRESA', 'PERFIL', 'FORMULARIO', 'PAINEL', 'ATIVO', 'OBSERVACAO'];
const USUARIO_DONO = '_dono'; // recuperação pelo código mestre, quando o dono não está na aba

function simNao_(v) { return v === true || semAcento_(texto_(v)).toUpperCase() === 'SIM'; }
function origemPlanilha_() {
  return SpreadsheetApp.openById(texto_(planilhaControle_().getSheetByName(CTRL.conf).getRange(CTRL.celulaOrigemId).getValue()));
}
/** Usuários da aba USUARIOS (guardados por 2 minutos; o painel limpa a cada alteração). */
function usuariosPainel_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('PAINEL_USUARIOS_V1');
  if (c) return JSON.parse(c);
  const sh = origemPlanilha_().getSheetByName(ABA_USUARIOS);
  const lista = [];
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach((r, i) => {
    const email = texto_(r[0]).toLowerCase();
    if (!email) return;
    lista.push({linha: i + 2, email, nome: texto_(r[1]), usuario: texto_(r[2]).toLowerCase(), empresa: texto_(r[3]), perfil: texto_(r[4]),
      formulario: simNao_(r[5]), painel: simNao_(r[6]), ativo: simNao_(r[7]), obs: texto_(r[8])});
  });
  try { cache.put('PAINEL_USUARIOS_V1', JSON.stringify(lista), 120); } catch (e) {}
  return lista;
}
function limparCacheUsuarios_() { try { CacheService.getScriptCache().remove('PAINEL_USUARIOS_V1'); } catch (e) {} }
function usuarioPorLogin_(login) {
  const u = texto_(login).toLowerCase();
  if (!u) return null;
  if (u === USUARIO_DONO) return {email: DONO_SISTEMA, nome: 'Dono do sistema', usuario: USUARIO_DONO, empresa: 'Zayda', perfil: 'Admin', painel: true, ativo: true};
  return usuariosPainel_().find(x => x.usuario === u) || null;
}
const podePainel_ = u => !!u && u.painel && u.ativo;
const ehAdmin_ = u => !!u && texto_(u.perfil).toUpperCase() === 'ADMIN';
const ehAdminZayda_ = u => ehAdmin_(u) && texto_(u.empresa).toUpperCase() === 'ZAYDA';
/** Outros fornecedores: quem pode conferir lançamentos e registrar pagamentos. Ativo e da Zayda (qualquer perfil) ou perfil Admin
 *  (hoje, a Mayara, financeiro da BSJ). A mesma regra está no formulário (Código.js, podeFinanceiro_). */
const podeFinanceiro_ = u => !!u && !!u.ativo && (texto_(u.empresa).toUpperCase() === 'ZAYDA' || ehAdmin_(u));
const publico_ = u => ({nome: u.nome || u.usuario, usuario: u.usuario, empresa: u.empresa, perfil: u.perfil, email: u.email,
  admin: ehAdmin_(u), adminZayda: ehAdminZayda_(u), financeiro: podeFinanceiro_(u), dono: u.usuario === USUARIO_DONO});

/* Senha guardada como "v2:<repetições>:<sal>:<hash>": SHA-256 de (sal + senha), repetido sobre o próprio hash até
 * completar as repetições. O cálculo é feito aqui mesmo, em JavaScript (10.000 chamadas ao Utilities deixariam o login lento).
 * PIN antigo ("<sal>:<hash>", um SHA-256 de PIN + sal) continua valendo no login; ao entrar, a pessoa cria a senha nova. */
const REPETICOES_SENHA = 10000;
const MSG_SENHA = 'A senha deve ter de 8 a 32 caracteres, com pelo menos uma letra e um número.';
const senhaValida_ = s => typeof s === 'string' && s === s.trim() && s.length >= 8 && s.length <= 32 &&
  /[A-Za-zÀ-ÿ]/.test(s) && /\d/.test(s) && /^[\x20-\x7EÀ-ÿ]+$/.test(s);
function utf8_(s) { const e = unescape(encodeURIComponent(String(s))), out = []; for (let i = 0; i < e.length; i++) out.push(e.charCodeAt(i)); return out; }
const SHA_K_ = (() => { const k = [], h = []; let n = 2;
  const primo = x => { for (let d = 2; d * d <= x; d++) if (x % d === 0) return false; return true; };
  while (k.length < 64) { if (primo(n)) { if (h.length < 8) h.push((Math.pow(n, 1 / 2) % 1) * 4294967296 | 0); k.push((Math.pow(n, 1 / 3) % 1) * 4294967296 | 0); } n++; }
  return {k, h}; })();
/** SHA-256 de uma lista de bytes (0–255); devolve 32 bytes. */
function sha256_(m) {
  const K = SHA_K_.k, l = m.length, n = ((l + 9 + 63) >> 6) << 4, w = new Array(n).fill(0), W = new Array(64);
  for (let i = 0; i < l; i++) w[i >> 2] |= m[i] << (24 - (i % 4) * 8);
  w[l >> 2] |= 0x80 << (24 - (l % 4) * 8);
  w[n - 1] = l * 8;
  const ror = (x, r) => (x >>> r) | (x << (32 - r));
  let [h0, h1, h2, h3, h4, h5, h6, h7] = SHA_K_.h;
  for (let j = 0; j < n; j += 16) {
    for (let t = 0; t < 64; t++) {
      if (t < 16) W[t] = w[j + t] | 0;
      else { const a = W[t - 15], b = W[t - 2];
        W[t] = (W[t - 16] + (ror(a, 7) ^ ror(a, 18) ^ (a >>> 3)) + W[t - 7] + (ror(b, 17) ^ ror(b, 19) ^ (b >>> 10))) | 0; }
    }
    let a = h0, b = h1, c = h2, d = h3, e = h4, f = h5, g = h6, h = h7;
    for (let t = 0; t < 64; t++) {
      const t1 = (h + (ror(e, 6) ^ ror(e, 11) ^ ror(e, 25)) + ((e & f) ^ (~e & g)) + K[t] + W[t]) | 0;
      const t2 = ((ror(a, 2) ^ ror(a, 13) ^ ror(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
    }
    h0 = (h0 + a) | 0; h1 = (h1 + b) | 0; h2 = (h2 + c) | 0; h3 = (h3 + d) | 0; h4 = (h4 + e) | 0; h5 = (h5 + f) | 0; h6 = (h6 + g) | 0; h7 = (h7 + h) | 0;
  }
  const out = [];
  [h0, h1, h2, h3, h4, h5, h6, h7].forEach(x => out.push((x >>> 24) & 255, (x >>> 16) & 255, (x >>> 8) & 255, x & 255));
  return out;
}
function hashSenha_(senha, sal, repeticoes) {
  let h = sha256_(utf8_(sal + senha));
  for (let i = 1; i < repeticoes; i++) h = sha256_(h);
  return h.map(b => ('0' + b.toString(16)).slice(-2)).join('');
}
/** PIN antigo (formato "sal:hash"). */
function hashPin_(pin, sal) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin) + sal, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function guardarSenha_(usuario, senha) {
  const sal = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('PIN_' + usuario, 'v2:' + REPETICOES_SENHA + ':' + sal + ':' + hashSenha_(senha, sal, REPETICOES_SENHA));
}
/** 'v2' (senha nova), 'antigo' (PIN antigo) ou '' (não confere). */
function senhaConfere_(usuario, senha) {
  const v = PropertiesService.getScriptProperties().getProperty('PIN_' + usuario);
  if (!v) return '';
  const p = v.split(':');
  if (p.length === 4 && p[0] === 'v2') { const n = Number(p[1]); return n > 0 && n <= 1000000 && hashSenha_(senha, p[2], n) === p[3] ? 'v2' : ''; }
  if (p.length === 2) return hashPin_(senha, p[0]) === p[1] ? 'antigo' : '';
  return '';
}
const temPin_ = usuario => !!PropertiesService.getScriptProperties().getProperty('PIN_' + usuario);
const textoSenha_ = s => typeof s === 'string' ? s : (s === null || s === undefined ? '' : String(s)); // sem trim: espaços contam

/** Chave do aparelho: aleatória e longa, guardada só no servidor e no aparelho. */
function novoToken_(usuario) {
  const t = (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const agora = Date.now();
  PropertiesService.getScriptProperties().setProperty('TOKEN_' + t, JSON.stringify({u: usuario, criado: agora, uso: agora}));
  return t;
}
function apagarTokens_(usuario) {
  const props = PropertiesService.getScriptProperties(), todas = props.getProperties();
  let n = 0;
  Object.keys(todas).forEach(k => {
    if (k.indexOf('TOKEN_') !== 0) return;
    try { if (JSON.parse(todas[k]).u === usuario) { props.deleteProperty(k); n++; } } catch (e) {}
  });
  return n;
}
function contarTokens_() {
  const todas = PropertiesService.getScriptProperties().getProperties(), out = {};
  Object.keys(todas).forEach(k => { if (k.indexOf('TOKEN_') === 0) try { const u = JSON.parse(todas[k]).u; out[u] = (out[u] || 0) + 1; } catch (e) {} });
  return out;
}
/** Valida a chave do aparelho e confere de novo, a cada chamada, se o usuário continua com PAINEL Sim e ATIVO Sim. */
function autenticar_(token) {
  const t = texto_(token);
  if (!/^[0-9a-f]{96}$/.test(t)) return {erro: {acesso: false, motivo: 'TOKEN'}};
  const props = PropertiesService.getScriptProperties(), v = props.getProperty('TOKEN_' + t);
  if (!v) return {erro: {acesso: false, motivo: 'TOKEN'}};
  let reg; try { reg = JSON.parse(v); } catch (e) { return {erro: {acesso: false, motivo: 'TOKEN'}}; }
  if (Date.now() - Number(reg.uso || reg.criado || 0) > TOKEN_VALIDADE_DIAS * 864e5) { props.deleteProperty('TOKEN_' + t); return {erro: {acesso: false, motivo: 'EXPIRADO'}}; } // [Etapa 1]
  const u = usuarioPorLogin_(reg.u);
  if (!podePainel_(u)) { props.deleteProperty('TOKEN_' + t); return {erro: {acesso: false, motivo: 'DESATIVADO'}}; }
  if (Date.now() - Number(reg.uso || 0) > 3600000) { reg.uso = Date.now(); props.setProperty('TOKEN_' + t, JSON.stringify(reg)); } // último uso, no máximo 1 gravação por hora
  marcarVisto_(u.usuario);
  return {u};
}
/** Presença no painel: horário no cache ("visto_<usuario>", 6 h) a cada chamada autenticada e,
 *  no máximo a cada 5 minutos por usuário, também em "ULTIMO_PAINEL_<usuario>" (sobrevive ao cache). */
function marcarVisto_(usuario) {
  try {
    const agora = String(Date.now()), cache = CacheService.getScriptCache();
    cache.put('visto_' + usuario, agora, 21600);
    if (!cache.get('vistoGravado_' + usuario)) {
      PropertiesService.getScriptProperties().setProperty('ULTIMO_PAINEL_' + usuario, agora);
      cache.put('vistoGravado_' + usuario, '1', 300);
    }
  } catch (e) {}
}
/** Último horário (ms) de cada e-mail na aba _PRESENCA da planilha do formulário (B = yyyy-MM-dd HH:mm:ss, fuso de São Paulo). */
function presencaFormulario_() {
  const out = {};
  try {
    const sh = origemPlanilha_().getSheetByName('_PRESENCA');
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
      const email = texto_(r[0]).toLowerCase(); if (!email) return;
      let ms = 0;
      if (r[1] instanceof Date) ms = r[1].getTime();
      else if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(texto_(r[1]))) ms = Utilities.parseDate(texto_(r[1]), CTRL.fuso, 'yyyy-MM-dd HH:mm:ss').getTime();
      if (ms > (out[email] || 0)) out[email] = ms;
    });
  } catch (e) { Logger.log('Presença do formulário não lida: ' + e.message); }
  return out;
}

/** Login com usuário + senha. 5 erros seguidos no mesmo usuário bloqueiam por 15 minutos.
 *  Quem ainda tem o PIN antigo entra com ele e recebe novaSenha: true (o painel pede a senha nova). */
function entrarPainel(usuario, senha) {
  const login = texto_(usuario).toLowerCase(), cache = CacheService.getScriptCache(), s = textoSenha_(senha);
  if (!login) return {acesso: false, motivo: 'INCORRETO'};
  if (cache.get('PINBLOQ_' + login)) return {acesso: false, motivo: 'BLOQUEADO_PIN'};
  const u = login === USUARIO_DONO ? null : usuarioPorLogin_(login);
  const erro = () => {
    const n = Number(cache.get('PINERR_' + login) || 0) + 1;
    if (n >= 5) { cache.put('PINBLOQ_' + login, '1', 900); cache.remove('PINERR_' + login); return {acesso: false, motivo: 'BLOQUEADO_PIN'}; }
    cache.put('PINERR_' + login, String(n), 900);
    Utilities.sleep(800);
    return {acesso: false, motivo: 'INCORRETO'};
  };
  if (u && !temPin_(u.usuario) && podePainel_(u)) return {acesso: false, motivo: 'SEM_PIN'};
  const confere = u && s && s.length <= 64 ? senhaConfere_(u.usuario, s) : '';
  if (!confere) return erro();
  if (!podePainel_(u)) return {acesso: false, motivo: 'DESATIVADO'};
  cache.remove('PINERR_' + login);
  return {acesso: true, token: novoToken_(u.usuario), novaSenha: confere === 'antigo'};
}

/** Primeiro acesso (passo 1): usuário válido, ainda sem senha, com o código de acesso do painel.
 *  Sem usuário, só o código mestre: recuperação do dono (entra direto). */
function verificarPrimeiroAcesso(usuario, codigo) {
  const login = texto_(usuario).toLowerCase();
  const ac = validarCodigoPainel_(codigo);
  if (!login) {
    if (ac !== 'OK') return {acesso: false, motivo: ac === 'INVALIDO' ? 'CODIGO_INCORRETO' : ac};
    const dono = usuariosPainel_().find(x => x.email === DONO_SISTEMA);
    return {acesso: true, token: novoToken_(dono && podePainel_(dono) ? dono.usuario : USUARIO_DONO)};
  }
  if (ac !== 'OK') return {acesso: false, motivo: ac === 'INVALIDO' ? 'CODIGO_INCORRETO' : ac};
  const u = usuarioPorLogin_(login);
  if (!u || login === USUARIO_DONO) return {acesso: false, motivo: 'CODIGO_INCORRETO'};
  if (!podePainel_(u)) return {acesso: false, motivo: 'DESATIVADO'};
  if (temPin_(u.usuario)) return {acesso: false, motivo: 'COM_PIN'};
  return {acesso: true, criarPin: true, nome: u.nome};
}
/** Primeiro acesso (passo 2): cria a senha e conecta o aparelho. */
function primeiroAcesso(usuario, codigo, senha) {
  const v = verificarPrimeiroAcesso(usuario, codigo);
  if (!v.acesso || !v.criarPin) return v.acesso ? {acesso: false, motivo: 'CODIGO_INCORRETO'} : v;
  const s = textoSenha_(senha);
  if (!senhaValida_(s)) return {acesso: false, motivo: 'SENHA_FORMATO'};
  const u = usuarioPorLogin_(usuario);
  guardarSenha_(u.usuario, s);
  return {acesso: true, token: novoToken_(u.usuario)};
}
/** Troca a senha (também usada para sair do PIN antigo: a senha atual pode ser o PIN). */
function trocarSenha(token, senhaAtual, senhaNova) {
  const a = autenticar_(token); if (a.erro) return a.erro;
  if (a.u.usuario === USUARIO_DONO) return {ok: false, motivo: 'A recuperação pelo código mestre não tem senha.'};
  const nova = textoSenha_(senhaNova);
  if (!senhaValida_(nova)) return {ok: false, motivo: MSG_SENHA};
  const cache = CacheService.getScriptCache(), login = a.u.usuario;
  if (cache.get('PINBLOQ_' + login)) return {ok: false, motivo: 'Muitas tentativas. Tente de novo em 15 minutos.'};
  if (!senhaConfere_(login, textoSenha_(senhaAtual))) {
    const n = Number(cache.get('PINERR_' + login) || 0) + 1;
    if (n >= 5) { cache.put('PINBLOQ_' + login, '1', 900); cache.remove('PINERR_' + login); return {ok: false, motivo: 'Muitas tentativas. Tente de novo em 15 minutos.'}; }
    cache.put('PINERR_' + login, String(n), 900);
    return {ok: false, motivo: 'A senha atual está incorreta.'};
  }
  cache.remove('PINERR_' + login);
  guardarSenha_(login, nova);
  return {ok: true};
}
/** Nome antigo, mantido para não quebrar chamadas já existentes. */
function trocarPin(token, pinAtual, pinNovo) { return trocarSenha(token, pinAtual, pinNovo); }
function sairAparelho(token) {
  const t = texto_(token);
  if (/^[0-9a-f]{96}$/.test(t)) PropertiesService.getScriptProperties().deleteProperty('TOKEN_' + t);
  return {ok: true};
}

/* ===================== GESTÃO DE USUÁRIOS (só Admin da Zayda) ===================== */
const SO_ADMIN_ZAYDA = 'Somente administradores da Zayda podem gerenciar usuários.';
function adminZayda_(token) {
  const a = autenticar_(token);
  if (a.erro) return a;
  if (!ehAdminZayda_(a.u)) return {erro: {ok: false, motivo: SO_ADMIN_ZAYDA}};
  return a;
}
function auditarUsuario_(quem, alvo, descricao) {
  try {
    const sh = origemPlanilha_().getSheetByName('AUDITORIA');
    if (sh) sh.appendRow([Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm:ss'), quem.email, 'USUARIO', alvo, '',
      descricao + ' (por ' + (quem.nome || quem.usuario) + ')']);
  } catch (e) { Logger.log('Auditoria de usuário não gravada: ' + e.message); }
}
function listarUsuarios(token) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const aparelhos = contarTokens_(), lista = usuariosPainel_();
  // Presença: painel (cache, senão propriedades) e formulário (aba _PRESENCA), em milissegundos; 0 = nunca.
  const logins = lista.map(u => u.usuario).filter(Boolean), props = PropertiesService.getScriptProperties().getProperties();
  let vistos = {}; try { vistos = logins.length ? CacheService.getScriptCache().getAll(logins.map(l => 'visto_' + l)) : {}; } catch (e) {}
  const form = presencaFormulario_();
  const ultimoPainel = l => !l ? 0 : Math.max(Number(vistos['visto_' + l] || 0), Number(props['ULTIMO_PAINEL_' + l] || 0)) || 0;
  return {ok: true, eu: a.u.email, dono: DONO_SISTEMA, agora: Date.now(), usuarios: lista.map(u => ({email: u.email, nome: u.nome, usuario: u.usuario,
    empresa: u.empresa, perfil: u.perfil, formulario: u.formulario, painel: u.painel, ativo: u.ativo, obs: u.obs,
    temPin: temPin_(u.usuario), aparelhos: aparelhos[u.usuario] || 0, ultimoPainel: ultimoPainel(u.usuario), ultimoFormulario: form[u.email] || 0}))};
}
/** Inclui ou altera uma linha da aba USUARIOS (gravada como texto) e ajusta o acesso de editor à planilha do formulário. */
function salvarUsuario(token, dados) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  dados = dados || {};
  const sn = v => (v === true || simNao_(v)) ? 'Sim' : 'Não';
  const novo = {email: texto_(dados.email).toLowerCase(), nome: texto_(dados.nome), usuario: texto_(dados.usuario).toLowerCase(),
    empresa: texto_(dados.empresa), perfil: texto_(dados.perfil), formulario: sn(dados.formulario), painel: sn(dados.painel),
    ativo: sn(dados.ativo), obs: texto_(dados.obs)};
  const original = texto_(dados.emailOriginal).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(novo.email)) return {ok: false, motivo: 'Informe um e-mail válido.'};
  if (!novo.nome) return {ok: false, motivo: 'Informe o nome.'};
  if (!/^[a-z0-9._-]{2,30}$/.test(novo.usuario)) return {ok: false, motivo: 'O usuário deve ter de 2 a 30 letras sem acento, números, ponto, hífen ou sublinhado.'};
  if (['Zayda', 'BSJ'].indexOf(novo.empresa) < 0) return {ok: false, motivo: 'Empresa deve ser Zayda ou BSJ.'};
  if (['Admin', 'Pedidos'].indexOf(novo.perfil) < 0) return {ok: false, motivo: 'Perfil deve ser Admin ou Pedidos.'};
  if ([novo.nome, novo.obs].some(t => /^[=+@-]/.test(t))) return {ok: false, motivo: 'Nome e observação não podem começar com =, +, - ou @.'};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return {ok: false, motivo: 'O sistema está ocupado. Tente de novo em alguns segundos.'};
  try {
    limparCacheUsuarios_();
    const lista = usuariosPainel_(), antes = original ? lista.find(x => x.email === original) : null;
    if (original && !antes) return {ok: false, motivo: 'Usuário não encontrado. Atualize a lista.'};
    if ((antes && antes.email === DONO_SISTEMA) || (!antes && novo.email === DONO_SISTEMA) || (antes && novo.email === DONO_SISTEMA && antes.email !== DONO_SISTEMA))
      return {ok: false, motivo: 'A linha do dono do sistema não pode ser alterada pelo painel.'};
    if (lista.some(x => x !== antes && x.email === novo.email)) return {ok: false, motivo: 'Já existe um usuário com este e-mail.'};
    if (lista.some(x => x !== antes && x.usuario === novo.usuario) || novo.usuario === USUARIO_DONO) return {ok: false, motivo: 'Já existe um usuário com este login.'};
    if (antes && antes.email === a.u.email && (novo.ativo === 'Não' || novo.painel === 'Não')) return {ok: false, motivo: 'Você não pode desativar o seu próprio acesso.'};
    const origem = origemPlanilha_(), sh = origem.getSheetByName(ABA_USUARIOS);
    if (!sh) return {ok: false, motivo: 'A aba USUARIOS não foi encontrada na planilha do formulário.'};
    const linha = antes ? antes.linha : Math.max(2, sh.getLastRow() + 1);
    const valores = [novo.email, novo.nome, novo.usuario, novo.empresa, novo.perfil, novo.formulario, novo.painel, novo.ativo, novo.obs];
    sh.getRange(linha, 1, 1, 9).setNumberFormat('@').setValues([valores]);
    SpreadsheetApp.flush();
    limparCacheUsuarios_();
    // Login trocado: a senha e os aparelhos acompanham.
    const props = PropertiesService.getScriptProperties();
    if (antes && antes.usuario && antes.usuario !== novo.usuario) {
      const pin = props.getProperty('PIN_' + antes.usuario);
      if (pin) { props.setProperty('PIN_' + novo.usuario, pin); props.deleteProperty('PIN_' + antes.usuario); }
      const todas = props.getProperties();
      Object.keys(todas).forEach(k => { if (k.indexOf('TOKEN_') === 0) try { const r = JSON.parse(todas[k]); if (r.u === antes.usuario) { r.u = novo.usuario; props.setProperty(k, JSON.stringify(r)); } } catch (e) {} });
    }
    let desconectados = 0;
    if (novo.ativo === 'Não' || novo.painel === 'Não') desconectados = apagarTokens_(novo.usuario);
    // Editor da planilha do formulário (o formulário roda como quem acessa). Nunca mexe no dono da planilha.
    let avisoEditor = '';
    const donoPlanilha = (() => { try { return texto_(origem.getOwner().getEmail()).toLowerCase(); } catch (e) { return ''; } })();
    const editores = () => origem.getEditors().map(x => texto_(x.getEmail()).toLowerCase());
    try {
      if (antes && antes.email !== novo.email && antes.email !== donoPlanilha && antes.email !== DONO_SISTEMA && editores().indexOf(antes.email) >= 0) origem.removeEditor(antes.email);
      const ed = editores();
      if (novo.formulario === 'Sim' && novo.ativo === 'Sim') { if (ed.indexOf(novo.email) < 0) origem.addEditor(novo.email); }
      else if (novo.email !== donoPlanilha && novo.email !== DONO_SISTEMA && ed.indexOf(novo.email) >= 0) origem.removeEditor(novo.email);
    } catch (e) { avisoEditor = ' Atenção: não foi possível ajustar o acesso à planilha do formulário (' + e.message + ').'; }
    const campos = ['email', 'nome', 'usuario', 'empresa', 'perfil', 'formulario', 'painel', 'ativo', 'obs'];
    const rot = {email: 'e-mail', nome: 'nome', usuario: 'usuário', empresa: 'empresa', perfil: 'perfil', formulario: 'formulário', painel: 'painel', ativo: 'ativo', obs: 'observação'};
    const sn2 = v => v === true ? 'Sim' : v === false ? 'Não' : v;
    const mud = antes ? campos.filter(c => texto_(sn2(antes[c])) !== texto_(novo[c])).map(c => rot[c] + ': ' + (texto_(sn2(antes[c])) || '—') + ' → ' + (novo[c] || '—')) : [];
    auditarUsuario_(a.u, novo.email, antes ? 'Usuário ' + novo.usuario + ' alterado: ' + (mud.join('; ') || 'sem mudança')
      : 'Usuário ' + novo.usuario + ' incluído (' + novo.empresa + ', ' + novo.perfil + ', formulário ' + novo.formulario + ', painel ' + novo.painel + ', ativo ' + novo.ativo + ')');
    return {ok: true, desconectados, aviso: avisoEditor};
  } catch (e) { return {ok: false, motivo: 'Não foi possível salvar: ' + e.message}; }
  finally { lock.releaseLock(); }
}
function alvoGestao_(a, usuario) {
  const alvo = usuariosPainel_().find(x => x.usuario === texto_(usuario).toLowerCase());
  if (!alvo) return {erro: {ok: false, motivo: 'Usuário não encontrado. Atualize a lista.'}};
  if (alvo.email === DONO_SISTEMA && a.u.email !== DONO_SISTEMA) return {erro: {ok: false, motivo: 'A linha do dono do sistema não pode ser alterada pelo painel.'}};
  return {alvo};
}
/** Apaga a senha e desconecta todos os aparelhos: a pessoa faz o primeiro acesso de novo. */
function resetarPin(token, usuario) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const g = alvoGestao_(a, usuario); if (g.erro) return g.erro;
  PropertiesService.getScriptProperties().deleteProperty('PIN_' + g.alvo.usuario);
  const n = apagarTokens_(g.alvo.usuario);
  auditarUsuario_(a.u, g.alvo.email, 'Senha do usuário ' + g.alvo.usuario + ' apagada; ' + n + ' aparelho(s) desconectado(s)');
  return {ok: true, desconectados: n};
}
function desconectarAparelhos(token, usuario) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const g = alvoGestao_(a, usuario); if (g.erro) return g.erro;
  const n = apagarTokens_(g.alvo.usuario);
  auditarUsuario_(a.u, g.alvo.email, n + ' aparelho(s) do usuário ' + g.alvo.usuario + ' desconectado(s)');
  return {ok: true, desconectados: n};
}

/** Chamado pelo Painel.html com a chave do aparelho. Devolve só valores simples (sem objetos Date). */
function getDadosPainel(token) {
  const ac = autenticar_(token); // chave do aparelho + usuário ainda com PAINEL Sim e ATIVO Sim
  if (ac.erro) return ac.erro;
  sincronizarSeHouverNovidade_();
  return dadosPainel_(ac.u);
}
/** Dados do painel para um usuário (quem = null: sem dados de administrador; usado pelo resumo semanal por e-mail). */
function dadosPainel_(quem) {
  const ss = planilhaControle_();
  const cat = {};
  linhas_(ss, CTRL.destCatalogo, 2, 3).forEach(r => { if (r[0]) cat[texto_(r[0])] = texto_(r[2]); });
  // Status da planilha de obras -> rótulos usados pelo painel.
  const status = s => ({'EM ANDAMENTO': 'EM EXECUÇÃO', 'A INICIAR': 'PLANEJADA', 'PARADA': 'PARALISADA',
                        'CONCLUÍDA': 'CONCLUÍDA', 'A VERIFICAR': 'A VERIFICAR'})[texto_(s).toUpperCase()] || texto_(s).toUpperCase();
  // Etapas de cada obra (aba _IMPORT_ETAPAS, copiada da planilha de obras).
  const shE = ss.getSheetByName(CTRL.destEtapas), etapas = {}; let nomesEtapas = [];
  if (shE && shE.getLastRow() > 1) {
    const v = shE.getRange(1, 1, shE.getLastRow(), shE.getLastColumn()).getValues();
    nomesEtapas = v[0].slice(2, v[0].length - 1).map(texto_);
    v.slice(1).forEach(r => { if (texto_(r[0])) etapas[texto_(r[0])] = {lista: r.slice(2, r.length - 1).map(texto_), pct: num_(r[r.length - 1])}; });
  }
  const obras = linhas_(ss, CTRL.obras, 5, 10).filter(r => texto_(r[0])).map(r => ({
    id: texto_(r[0]), empreendimento: texto_(r[1]), casa: texto_(r[2]) || texto_(r[0]),
    cidade: texto_(r[4]), responsavel: '', status: status(r[7]), inicio: dataIso_(r[8]),
    cliente: texto_(r[1]), endereco: texto_(r[2]), fase: texto_(r[4]), prazo: texto_(r[5]),
    entrega: dataIso_(r[6]), statusTxt: texto_(r[7]), contrato: dataIso_(r[8]),
    pct: (etapas[texto_(r[0])] || {}).pct, etapas: (etapas[texto_(r[0])] || {}).lista || []
  }));
  const itens = linhas_(ss, CTRL.solicitacoes, 5, 18).filter(r => texto_(r[0]) && texto_(r[2])).map(r => ({
    pedido: texto_(r[0]), rev: num_(r[1]), sku: texto_(r[2]), material: texto_(r[3]), un: texto_(r[4]),
    qtdSol: num_(r[5]) || 0, precoRef: num_(r[6]), valorRef: num_(r[7]), data: dataIso_(r[8]),
    obra: texto_(r[9]), entregue: num_(r[11]) || 0, saldo: num_(r[12]) || 0, situacao: texto_(r[13]),
    obraTexto: texto_(r[14]), atendente: texto_(r[15]), entregador: texto_(r[16]), statusPedido: texto_(r[17])
  }));
  const lanc = linhas_(ss, CTRL.lancamentos, 5, 16).filter(r => texto_(r[0])).map(r => ({
    id: texto_(r[0]), data: dataIso_(r[1]), obra: texto_(r[2]), tipo: texto_(r[3]), pedido: texto_(r[4]),
    sku: texto_(r[5]), material: texto_(r[6]), cat: cat[texto_(r[5])] || 'Outros', un: texto_(r[7]),
    qtd: num_(r[8]) || 0, custo: num_(r[9]), valor: num_(r[10]), resp: texto_(r[11]), conf: texto_(r[12]),
    obs: texto_(r[13]), situacao: texto_(r[14]), alertas: texto_(r[15])
  }));
  // Dados de cabeçalho de cada pedido (revisão vigente): usados na impressão do pedido.
  const pedidos = {}, primeiraRev = {}; // primeiraRev: e-mail de quem criou o pedido (menor revisão)
  linhas_(ss, CTRL.destPedidos, 2, 17).filter(r => texto_(r[0])).forEach(r => {
    const id = texto_(r[0]), rev = Number(r[1]) || 0;
    if (!primeiraRev[id] || rev < primeiraRev[id].rev) primeiraRev[id] = {rev, email: texto_(r[4]).toLowerCase()};
    if (pedidos[id] && pedidos[id].rev > rev) return;
    pedidos[id] = {rev, criado: dataHora_(r[2]), atualizado: dataHora_(r[3]), email: texto_(r[4]), atendente: texto_(r[5]),
      entregador: texto_(r[6]), endereco: texto_(r[7]), casa: texto_(r[8]), obs: texto_(r[9]), status: texto_(r[12]),
      referencia: texto_(r[15]), etapa: texto_(r[16])}; // [Função 6] Q: etapa da revisão vigente
  });
  // Entregas de cada pedido, com quem entregou e quem recebeu (_IMPORT_ENTREGAS), da mais antiga para a mais nova.
  Object.keys(pedidos).forEach(id => { pedidos[id].criadoPor = primeiraRev[id] ? primeiraRev[id].email : ''; });
  const shEnt = ss.getSheetByName(CTRL.destEntregas);
  if (shEnt && shEnt.getLastRow() > 1) shEnt.getRange(2, 1, shEnt.getLastRow() - 1, Math.min(11, shEnt.getMaxColumns())).getValues().forEach(r => {
    const id = texto_(r[1]);
    if (!pedidos[id] || !texto_(r[0])) return;
    (pedidos[id].entregas = pedidos[id].entregas || []).push({id: texto_(r[0]), sku: texto_(r[2]), material: texto_(r[3]), un: texto_(r[4]),
      qtd: num_(r[5]) || 0, data: dataIso_(r[6]), recebidoPor: texto_(r[7]), entreguePor: texto_(r[8]), obs: texto_(r[9]),
      registradoPor: texto_(r[10]).toLowerCase()});
  });
  Object.keys(pedidos).forEach(id => { if (pedidos[id].entregas) pedidos[id].entregas.sort((a, b) => a.data.localeCompare(b.data) || a.id.localeCompare(b.id)); });
  // Pedido cancelado ou encerrado no formulário (_IMPORT_CANCELAMENTOS). Vale o primeiro registro de cada pedido.
  const shC = ss.getSheetByName(CTRL.destCancel);
  if (shC && shC.getLastRow() > 1) shC.getRange(2, 1, shC.getLastRow() - 1, 6).getValues().forEach(r => {
    const id = texto_(r[0]);
    if (!pedidos[id] || pedidos[id].cancelamento) return;
    pedidos[id].cancelamento = {tipo: texto_(r[1]).toUpperCase() === 'ENCERRADO' ? 'ENCERRADO' : 'CANCELADO', data: dataHora_(r[2]),
      email: texto_(r[3]), motivo: texto_(r[4]), naoEntregues: texto_(r[5])};
  });
  // Orçamentos preenchidos na aba ORCAMENTOS (só as obras com valor).
  const orcamentos = {}, shOrc = ss.getSheetByName(CTRL.orcamentos);
  if (shOrc && shOrc.getLastRow() > 1) {
    shOrc.getRange(2, 1, shOrc.getLastRow() - 1, 6).getValues().forEach(r => {
      const id = texto_(r[0]).toUpperCase(), valor = valor_(r[3]);
      if (id && valor > 0 && !orcamentos[id]) orcamentos[id] = {valor,
        data: r[4] instanceof Date ? Utilities.formatDate(r[4], CTRL.fuso, 'dd/MM/yyyy') : texto_(r[4]), obs: texto_(r[5])};
    });
  }
  // Outros fornecedores (planilha do formulário). Uma falha aqui não impede o painel de abrir.
  let ter = {terceiros: [], fornecedores: {}};
  try { ter = lerTerceiros_(); } catch (e) { Logger.log('Outros fornecedores não lidos: ' + e.message); }
  const conf = ss.getSheetByName(CTRL.conf);
  return {
    acesso: true, pedidos, nomesEtapas, orcamentos, terceiros: ter.terceiros, fornecedores: ter.fornecedores, formasPag: TERC.formas,
    // Dados técnicos (sincronização, backup e link da planilha): só vão para Admin Zayda; para os outros nem saem do servidor.
    sistema: ehAdminZayda_(quem) ? {sincErro: (() => { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('SINC_ERRO') || 'null'); } catch (e) { return null; } })(), sinc: resumoSinc_(), backup: (() => { try { const b = JSON.parse(PropertiesService.getScriptProperties().getProperty('BACKUP_ULTIMO') || 'null'); if (b && !b.ok) return {ok: false, em: b.em, origem: b.origem, falta: b.falta || (/permiss|permission|DriveApp/i.test(String(b.erro || '')) ? 'AUTORIZAR' : 'OUTRO')}; return b; } catch (e) { return null; } })()} : null, // [Etapa 1]
    usuarios: usuariosPainel_().reduce((m, u) => { m[u.email] = {nome: u.nome, empresa: u.empresa}; return m; }, {}), eu: quem ? publico_(quem) : null,
    geradoEm: Utilities.formatDate(new Date(), CTRL.fuso, "yyyy-MM-dd'T'HH:mm:ss"),
    hoje: Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd'),
    modo: texto_(conf.getRange(CTRL.celulaModo).getValue()) || 'PRODUÇÃO',
    sincronizadoEm: texto_(conf.getRange(CTRL.celulaUltimaSinc).getValue()).split(' · ')[0],
    limitesOrc: ALERTA_ORC, // [Função 1]
    areas: areasObras_(), // [Função 3] m² de cada obra
    fechamentos: (() => { try { return lerFechamentos_(); } catch (e) { return []; } })(), // [Função 4] meses fechados (cadeado e conferência)
    resumo: ehAdminZayda_(quem) ? {destinos: destinosResumo_(), ultimo: ultimoResumo_(), hora: RESUMO.hora} : null, // [Função 5] só Admin Zayda
    planilhaUrl: ehAdminZayda_(quem) ? ss.getUrl() : '', versaoPainel: versaoPainel_(), obras, itens, lanc
  };
}

/** Entregas de um pedido, com nome e assinatura de quem recebeu (lidas da aba ENTREGAS do formulário).
 *  Chamado pelo painel só na hora de imprimir o pedido, para não pesar a carga do painel. */
function getEntregasPedido(token, pedido) {
  const ac = autenticar_(token);
  if (ac.erro) return ac.erro;
  const id = texto_(pedido).toUpperCase();
  const conf = planilhaControle_().getSheetByName(CTRL.conf);
  const sh = SpreadsheetApp.openById(texto_(conf.getRange(CTRL.celulaOrigemId).getValue())).getSheetByName(CTRL.abaEntregasOrigem);
  const n = sh ? sh.getLastRow() : 0;
  const entregas = n < 2 ? [] : sh.getRange(2, 1, n - 1, 13).getValues()
    .filter(r => texto_(r[1]).toUpperCase() === id && texto_(r[0]).indexOf(CTRL.prefixoEntrega) === 0)
    .map(r => ({id: texto_(r[0]), sku: texto_(r[3]), produto: texto_(r[4]), un: texto_(r[5]), qtd: Number(valor_(r[6])) || 0,
                data: dataIso_(r[7]) || dataHora_(r[7]).slice(0, 10), registrado: dataHora_(r[9]), obs: texto_(r[10]),
                recebido: texto_(r[11]), assinatura: /^data:image\/png;base64,/.test(texto_(r[12])) ? texto_(r[12]) : ''}));
  return {acesso: true, pedido: id, entregas};
}

/* ===================== OUTROS FORNECEDORES (lançados no formulário) =====================
 * Lidos direto da planilha do formulário na mesma chamada do painel (getDadosPainel), sem cópia na planilha de controle.
 * FORNECEDORES  A ID · B NOME · C TIPO_PRINCIPAL · D CPF_CNPJ · E TELEFONE · F OBSERVACAO · G ATIVO · H CRIADO_POR · I CRIADO_EM
 * TERCEIROS     A ID · B CRIADO_EM · C CRIADO_POR · D OBRA_ID · E FORNECEDOR_ID · F FORNECEDOR_NOME · G TIPO · H DESCRICAO · I QUANTIDADE
 *               J UNIDADE · K VALOR_TOTAL · L DATA_REFERENCIA · M NOTA_NUMERO · N ANEXO_URL · O FORMA_PAGAMENTO · P CONFERENCIA
 *               Q CONFERIDO_POR · R CONFERIDO_EM · S MOTIVO · T OBSERVACAO · U ATUALIZADO_EM · V ATUALIZADO_POR
 * TERCEIROS_PARCELAS A ID_PARCELA · B ID_LANCAMENTO · C NUMERO · D DESCRICAO · E VALOR · F VENCIMENTO · G PAGO · H DATA_PAGAMENTO
 *               I FORMA · J COMPROVANTE_URL · K REGISTRADO_POR · L REGISTRADO_EM
 * Só CONFERENCIA = CONFERIDO vira custo da obra (no painel). Abas ausentes: listas vazias. */
const TERC = Object.freeze({abaForn: 'FORNECEDORES', abaLanc: 'TERCEIROS', abaParc: 'TERCEIROS_PARCELAS', nLanc: 22, nParc: 12, nForn: 9,
  formas: ['Pix', 'Boleto', 'Cartão', 'Dinheiro', 'Transferência', 'Outro']}); // formas: a mesma lista do formulário (TER.formas em Código.js)
function linhasOrigem_(origem, nome, n) {
  const sh = origem.getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(n, sh.getMaxColumns())).getValues();
}
const urlDrive_ = v => { const s = texto_(v); return /^https:\/\/(drive|docs)\.google\.com\//.test(s) ? s : ''; };
function lerTerceiros_(origem) {
  origem = origem || origemPlanilha_();
  const fornecedores = {};
  linhasOrigem_(origem, TERC.abaForn, TERC.nForn).forEach(r => { const id = texto_(r[0]); if (/^FOR-\d+$/.test(id)) fornecedores[id] = {
    nome: textoForm_(r[1]), tipo: textoForm_(r[2]), doc: textoForm_(r[3]), telefone: textoForm_(r[4]), obs: textoForm_(r[5]), ativo: texto_(r[6]).toUpperCase() !== 'NAO'}; });
  const parc = {};
  linhasOrigem_(origem, TERC.abaParc, TERC.nParc).forEach(r => {
    const lanc = texto_(r[1]); if (texto_(r[0]).indexOf('TER-') !== 0 || !lanc) return;
    (parc[lanc] = parc[lanc] || []).push({id: texto_(r[0]), numero: Number(r[2]) || 0, descricao: textoForm_(r[3]), valor: Math.round((Number(valor_(r[4])) || 0) * 100) / 100,
      vencimento: dataIso_(r[5]), pago: texto_(r[6]).toUpperCase() === 'SIM', dataPagamento: dataIso_(r[7]), forma: textoForm_(r[8]),
      comprovante: urlDrive_(r[9]), registradoPor: texto_(r[10]).toLowerCase(), registradoEm: dataHora_(r[11])});
  });
  // [Função 6] ETAPA procurada pelo nome no cabeçalho da aba TERCEIROS (coluna nova no fim); sem a coluna, etapa vazia.
  const shTer = origem.getSheetByName(TERC.abaLanc), cabTer = shTer && shTer.getLastColumn() ? shTer.getRange(1, 1, 1, Math.min(shTer.getLastColumn(), 60)).getValues()[0].map(v => texto_(v).toUpperCase()) : [];
  const iEtapaT = cabTer.indexOf('ETAPA');
  const lista = linhasOrigem_(origem, TERC.abaLanc, Math.max(TERC.nLanc, iEtapaT + 1)).filter(r => texto_(r[0]).indexOf('TER-') === 0).map(r => {
    const id = texto_(r[0]);
    return {id, criadoEm: dataHora_(r[1]), criadoPor: texto_(r[2]).toLowerCase(), obra: texto_(r[3]).toUpperCase(), fornecedorId: texto_(r[4]),
      fornecedor: textoForm_(r[5]), tipo: textoForm_(r[6]) || 'Outros', descricao: textoForm_(r[7]), quantidade: num_(r[8]), unidade: textoForm_(r[9]),
      valor: Math.round((Number(valor_(r[10])) || 0) * 100) / 100, dataRef: dataIso_(r[11]), nota: textoForm_(r[12]), anexo: urlDrive_(r[13]),
      forma: texto_(r[14]).toUpperCase(), conferencia: texto_(r[15]).toUpperCase() || 'AGUARDANDO', conferidoPor: texto_(r[16]).toLowerCase(),
      conferidoEm: dataHora_(r[17]), motivo: textoForm_(r[18]), obs: textoForm_(r[19]), atualizadoEm: dataHora_(r[20]), atualizadoPor: texto_(r[21]).toLowerCase(),
      parcelas: (parc[id] || []).sort((a, b) => a.numero - b.numero), etapa: iEtapaT >= 0 ? textoForm_(r[iEtapaT]) : ''};
  });
  return {terceiros: lista, fornecedores};
}
/** Conferência no painel (só Admin da Zayda): op = CONFERIR | REPROVAR (AGUARDANDO) ou CANCELAR (CONFERIDO).
 *  Reprovar e cancelar exigem motivo. Grava na aba TERCEIROS e na AUDITORIA da planilha do formulário e devolve os dados novos. */
function conferirTerceiro(token, id, op, motivo, atualizadoEm) {
  const ac = autenticar_(token); if (ac.erro) return ac.erro;
  if (!podeFinanceiro_(ac.u)) return {ok: false, motivo: 'Só a equipe Zayda ou o financeiro podem conferir lançamentos.'};
  const acao = texto_(op).toUpperCase(), alvo = texto_(id);
  if (['CONFERIR', 'REPROVAR', 'CANCELAR'].indexOf(acao) < 0 || !/^TER-\d{8}-[A-Z0-9]{4}$/.test(alvo)) return {ok: false, motivo: 'Pedido inválido.'};
  const mot = texto_(motivo);
  if (acao !== 'CONFERIR') {
    if (mot.length < 5) return {ok: false, motivo: 'Informe o motivo (mínimo 5 caracteres).'};
    if (mot.length > 300) return {ok: false, motivo: 'Motivo muito longo (máximo 300 caracteres).'};
    if (/^[=+@\-]/.test(mot)) return {ok: false, motivo: 'O motivo não pode começar com =, +, - ou @.'};
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return {ok: false, motivo: 'Sistema ocupado. Tente de novo em instantes.'};
  try {
    const origem = origemPlanilha_(), sh = origem.getSheetByName(TERC.abaLanc);
    if (!sh || sh.getLastRow() < 2) return {ok: false, motivo: 'Lançamento não encontrado.'};
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => texto_(r[0])), i = ids.indexOf(alvo);
    if (i < 0) return {ok: false, motivo: 'Lançamento não encontrado.'};
    const linha = i + 2, r = sh.getRange(linha, 1, 1, TERC.nLanc).getValues()[0], atual = texto_(r[15]).toUpperCase() || 'AGUARDANDO';
    if (texto_(atualizadoEm) !== dataHora_(r[20])) return {ok: false, motivo: 'Este lançamento foi alterado no formulário depois que você abriu. A tela foi atualizada: confira de novo.', atualizar: true};
    const precisa = acao === 'CANCELAR' ? 'CONFERIDO' : 'AGUARDANDO';
    if (atual !== precisa) return {ok: false, motivo: 'Este lançamento está ' + atual.toLowerCase() + ': ' + (acao === 'CANCELAR' ? 'só um lançamento conferido pode ser cancelado.' : 'só um lançamento aguardando conferência pode ser conferido ou reprovado.'), atualizar: true};
    const novo = {CONFERIR: 'CONFERIDO', REPROVAR: 'REPROVADO', CANCELAR: 'CANCELADO'}[acao];
    const agora = Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm:ss');
    sh.getRange(linha, 16, 1, 4).setNumberFormat('@').setValues([[novo, ac.u.email, agora, acao === 'CONFERIR' ? '' : mot]]);
    const valor = (Number(valor_(r[10])) || 0).toFixed(2).replace('.', ',');
    const aud = origem.getSheetByName('AUDITORIA');
    if (aud) aud.appendRow([agora, ac.u.email, 'TERCEIRO_' + novo, alvo, '', 'Painel: ' + {CONFERIR: 'conferido', REPROVAR: 'reprovado', CANCELAR: 'cancelado'}[acao] +
      ' por ' + (ac.u.nome || ac.u.usuario) + ' · ' + texto_(r[5]) + ' · ' + texto_(r[6]) + ' · R$ ' + valor + ' · obra ' + texto_(r[3]) +
      (acao === 'CANCELAR' ? ' · conferido antes por ' + texto_(r[16]) + ' em ' + dataHora_(r[17]) : '') + (mot ? ' · Motivo: ' + mot : '')]);
    SpreadsheetApp.flush();
    const dados = lerTerceiros_(origem);
    return {ok: true, status: novo, terceiros: dados.terceiros, fornecedores: dados.fornecedores};
  } finally { lock.releaseLock(); }
}

/** Pagamento de parcela no painel (mesmo caminho da conferência). op = PAGAR (data aaaa-mm-dd até hoje + forma) ou DESFAZER (motivo).
 *  Grava na aba TERCEIROS_PARCELAS da planilha do formulário, nas mesmas colunas G–L que o formulário usa, e na AUDITORIA.
 *  Não mexe no custo da obra (que depende só da conferência). Devolve os dados novos. */
function pagarParcelaTerceiro(token, idParcela, op, data, forma, motivo) {
  const ac = autenticar_(token); if (ac.erro) return ac.erro;
  if (!podeFinanceiro_(ac.u)) return {ok: false, motivo: 'Só a equipe Zayda ou o financeiro podem registrar pagamentos.'};
  const acao = texto_(op).toUpperCase(), alvo = texto_(idParcela);
  if (['PAGAR', 'DESFAZER'].indexOf(acao) < 0 || !/^TER-\d{8}-[A-Z0-9]{4}-P\d{1,2}$/.test(alvo)) return {ok: false, motivo: 'Pedido inválido.'};
  const dia = texto_(data), fp = texto_(forma), mot = texto_(motivo);
  if (acao === 'PAGAR') {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dia), d = m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
    if (!d || d.getUTCDate() !== +m[3] || dia < '2020-01-01') return {ok: false, motivo: 'Data do pagamento inválida.'};
    if (TERC.formas.indexOf(fp) < 0) return {ok: false, motivo: 'Escolha a forma de pagamento.'};
  } else {
    if (mot.length < 5) return {ok: false, motivo: 'Informe o motivo (mínimo 5 caracteres).'};
    if (mot.length > 300) return {ok: false, motivo: 'Motivo muito longo (máximo 300 caracteres).'};
    if (/^[=+@\-]/.test(mot)) return {ok: false, motivo: 'O motivo não pode começar com =, +, - ou @.'};
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return {ok: false, motivo: 'Sistema ocupado. Tente de novo em instantes.'};
  try {
    const origem = origemPlanilha_(), sh = origem.getSheetByName(TERC.abaParc), shL = origem.getSheetByName(TERC.abaLanc);
    if (!sh || sh.getLastRow() < 2 || !shL || shL.getLastRow() < 2) return {ok: false, motivo: 'Parcela não encontrada.'};
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => texto_(r[0])), i = ids.indexOf(alvo);
    if (i < 0) return {ok: false, motivo: 'Parcela não encontrada.'};
    const linha = i + 2, r = sh.getRange(linha, 1, 1, TERC.nParc).getValues()[0], pago = texto_(r[6]).toUpperCase() === 'SIM';
    const lids = shL.getRange(2, 1, shL.getLastRow() - 1, 1).getValues().map(x => texto_(x[0])), j = lids.indexOf(texto_(r[1]));
    if (j < 0) return {ok: false, motivo: 'Lançamento da parcela não encontrado.'};
    const l = shL.getRange(j + 2, 1, 1, TERC.nLanc).getValues()[0];
    if ((texto_(l[15]).toUpperCase() || 'AGUARDANDO') === 'CANCELADO') return {ok: false, motivo: 'Este lançamento foi cancelado.'};
    if (acao === 'PAGAR' && pago) return {ok: false, motivo: 'Esta parcela já foi paga.'};
    if (acao === 'DESFAZER' && !pago) return {ok: false, motivo: 'Esta parcela não está paga.'};
    const agora = Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm:ss'), quem = ac.u.usuario || ac.u.nome || ac.u.email;
    const valor = (Number(valor_(r[4])) || 0).toFixed(2).replace('.', ',');
    // G PAGO · H DATA_PAGAMENTO · I FORMA · J COMPROVANTE_URL · K REGISTRADO_POR · L REGISTRADO_EM (comprovante só pelo formulário)
    sh.getRange(linha, 7, 1, 6).setNumberFormat('@').setValues([acao === 'PAGAR' ? ['SIM', dia, fp, texto_(r[9]), quem, agora] : ['NAO', '', '', '', '', '']]);
    const aud = origem.getSheetByName('AUDITORIA');
    if (aud) aud.appendRow([agora, ac.u.email, acao === 'PAGAR' ? 'TERCEIRO_PAGAMENTO' : 'TERCEIRO_PAGAMENTO_DESFEITO', texto_(r[1]), '', 'Painel: ' +
      (acao === 'PAGAR' ? 'parcela ' + alvo + ' (' + texto_(r[3]) + ', R$ ' + valor + ') paga em ' + dia + ' por ' + fp + ' · registrado por ' + (ac.u.nome || quem)
        : 'pagamento desfeito na parcela ' + alvo + ' (R$ ' + valor + ', pago em ' + dataIso_(r[7]) + ' por ' + texto_(r[8]) + ', registrado por ' + texto_(r[10]) +
          ' em ' + dataHora_(r[11]) + ') por ' + (ac.u.nome || quem) + ' · Motivo: ' + mot)]);
    SpreadsheetApp.flush();
    const dados = lerTerceiros_(origem);
    return {ok: true, status: acao === 'PAGAR' ? 'PAGO' : 'DESFEITO', terceiros: dados.terceiros, fornecedores: dados.fornecedores};
  } finally { lock.releaseLock(); }
}

/* ===================== [Etapa 1] ESTABILIDADE =====================
 * Medição da sincronização automática (cota de 90 min/dia de gatilhos), backup semanal e validade das chaves de aparelho. */
const SINC_DIAS = 14, SINC_COTA_MIN = 90;
/** [Função 1] Alerta de orçamento: % do orçamento já gasto (custo confirmado total da obra, desde o início).
 *  Abaixo de "atencao" = Dentro do orçamento · de "atencao" até menos de "estouro" = Atenção · "estouro" ou mais = Estourado. */
const ALERTA_ORC = Object.freeze({atencao: 0.8, estouro: 1});
/** Acumula a execução do dia (fuso São Paulo): execuções, milissegundos, maior duração e quantas foram "sem mudança". Guarda 14 dias. */
function registrarExecucaoSinc_(ms, semMudanca) {
  try {
    const props = PropertiesService.getScriptProperties(), dia = Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd'), k = 'SINC_DIA_' + dia;
    let x = {n: 0, ms: 0, max: 0, sem: 0}; try { x = JSON.parse(props.getProperty(k) || 'null') || x; } catch (e) {}
    x.n++; x.ms += ms; x.max = Math.max(x.max, ms); if (semMudanca) x.sem++;
    props.setProperty(k, JSON.stringify(x));
    if (x.n === 1) { // primeira do dia: apaga os dias com mais de 14 dias
      const limite = Utilities.formatDate(new Date(Date.now() - SINC_DIAS * 864e5), CTRL.fuso, 'yyyy-MM-dd');
      Object.keys(props.getProperties()).forEach(c => { if (c.indexOf('SINC_DIA_') === 0 && c.slice(9) < limite) props.deleteProperty(c); });
    }
  } catch (e) { Logger.log('Medição da sincronização não gravada: ' + e.message); }
}
/** Última falha da sincronização automática (só a data; o texto técnico fica no registro de execuções). Uma execução certa apaga o aviso. */
function marcarErroSinc_(msg) {
  try {
    const p = PropertiesService.getScriptProperties();
    if (!msg) { if (p.getProperty('SINC_ERRO')) p.deleteProperty('SINC_ERRO'); return; }
    p.setProperty('SINC_ERRO', JSON.stringify({em: Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm:ss')}));
    console.error('Sincronização automática falhou: ' + msg);
  } catch (e) {}
}
function resumoSinc_() {
  const props = PropertiesService.getScriptProperties(), dia = Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd');
  let x = null; try { x = JSON.parse(props.getProperty('SINC_DIA_' + dia) || 'null'); } catch (e) {}
  x = x || {n: 0, ms: 0, max: 0, sem: 0};
  return {dia, n: x.n, min: Math.round(x.ms / 600) / 100, media: x.n ? Math.round(x.ms / x.n / 100) / 10 : 0, max: Math.round(x.max / 100) / 10, sem: x.sem, cota: SINC_COTA_MIN};
}

/* Backup semanal: cópia das planilhas do FORMULÁRIO e do CONTROLE na pasta "ZAYDA - Backups" (Drive do dono), as 8 mais recentes de cada.
 * A cópia de uma planilha leva junto o projeto do Apps Script dela (os gatilhos não vão junto). Nada é compartilhado. */
const BACKUP = Object.freeze({pasta: 'ZAYDA - Backups', manter: 8, idFormulario: '1F7q4M4Sv1WRAY_9HNrPBmiCWymkCKS-EdAa6644GB6Q'});
function pastaBackup_() {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('BACKUP_PASTA_ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const it = DriveApp.getFoldersByName(BACKUP.pasta);
  const pasta = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP.pasta);
  props.setProperty('BACKUP_PASTA_ID', pasta.getId());
  return pasta;
}
function fazerBackup_(origem) {
  const props = PropertiesService.getScriptProperties(), inicio = Date.now();
  const agora = Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd HH:mm:ss'), dia = agora.slice(0, 10);
  try {
    const pasta = pastaBackup_();
    let idForm = '';
    try { idForm = texto_(planilhaControle_().getSheetByName(CTRL.conf).getRange(CTRL.celulaOrigemId).getValue()); } catch (e) {}
    const copias = [[idForm || BACKUP.idFormulario, 'Formulário'], [planilhaControle_().getId(), 'Controle']];
    copias.forEach(([id, nome]) => DriveApp.getFileById(id).makeCopy('BACKUP ' + dia + ' — ' + nome, pasta));
    // Mantém as 8 mais recentes de cada: só arquivos desta pasta cujo nome começa com "BACKUP ".
    const lixo = [];
    copias.forEach(([, nome]) => {
      const lista = [], it = pasta.getFiles();
      while (it.hasNext()) { const f = it.next(), n = f.getName(); if (n.indexOf('BACKUP ') === 0 && n.slice(-(' — ' + nome).length) === ' — ' + nome) lista.push(f); }
      lista.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(BACKUP.manter).forEach(f => { f.setTrashed(true); lixo.push(f.getName()); });
    });
    const tokens = limparTokensVelhos_();
    const r = {ok: true, em: agora, origem, segundos: Math.round((Date.now() - inicio) / 1000), lixeira: lixo.length, tokensRemovidos: tokens};
    props.setProperty('BACKUP_ULTIMO', JSON.stringify(r));
    return r;
  } catch (e) {
    // O texto técnico do Google fica só no registro de execuções; o painel recebe apenas o tipo da falha.
    const erro = String(e && e.message || e);
    const r = {ok: false, em: agora, origem, falta: /permiss|permission|DriveApp/i.test(erro) ? 'AUTORIZAR' : 'OUTRO'};
    props.setProperty('BACKUP_ULTIMO', JSON.stringify(r));
    console.error('Backup não feito: ' + erro);
    return r;
  }
}
/** Gatilho semanal (segunda, 3h–4h). */
function backupSemanal() { return fazerBackup_('semanal'); }
/** Gatilho de uma vez, criado pelo botão "Fazer backup agora" do painel (assim o painel não espera a cópia terminar). */
function backupAgendado() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backupAgendado').forEach(t => ScriptApp.deleteTrigger(t));
  return fazerBackup_('manual');
}
function backupAgora(token) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const pend = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupAgendado');
  if (!pend) ScriptApp.newTrigger('backupAgendado').timeBased().after(5000).create();
  return {ok: true, agendado: true, jaAgendado: pend};
}
/** Chaves de aparelho sem uso há mais de 60 dias deixam de valer. */
const TOKEN_VALIDADE_DIAS = 60;
function limparTokensVelhos_() {
  const props = PropertiesService.getScriptProperties(), todas = props.getProperties(), limite = Date.now() - TOKEN_VALIDADE_DIAS * 864e5;
  let n = 0;
  Object.keys(todas).forEach(k => {
    if (k.indexOf('TOKEN_') !== 0) return;
    try { const r = JSON.parse(todas[k]); if (Number(r.uso || r.criado || 0) < limite) { props.deleteProperty(k); n++; } } catch (e) {}
  });
  return n;
}

