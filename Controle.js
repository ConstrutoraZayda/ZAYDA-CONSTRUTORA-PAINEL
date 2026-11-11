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
  intervaloMinutos: 1, copiaCompletaMinutos: 15
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
  const pedidos = valores(aba(CTRL.celulaAbaPedidos), 15)
    .filter(valido)
    .map(r => { const p = [texto_(r[0]), Number(r[1]) || 0, dataHora_(r[2]), dataHora_(r[3]), texto_(r[4]), texto_(r[5]),
               texto_(r[6]), texto_(r[7]), texto_(r[8]), texto_(r[9]), valor_(r[10]),
               Number(valor_(r[11])) || 0, texto_(r[12]), texto_(r[13])]; p.referencia = texto_(r[14]); return p; });
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
 *  Chamado a cada minuto, a cada consulta do painel e na hora em que alguém edita a planilha de obras. */
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
  const cab = ['ENTREGA_ID', 'PEDIDO_ID', 'SKU', 'PRODUTO', 'UNIDADE', 'QTD_ENTREGUE', 'DATA_ENTREGA', 'RECEBIDO_POR', 'ENTREGUE_POR', 'OBSERVACOES'];
  let sh = ss.getSheetByName(CTRL.destEntregas);
  if (!sh) { sh = ss.insertSheet(CTRL.destEntregas, ss.getNumSheets()); sh.hideSheet(); }
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  const ultima = sh.getMaxRows();
  if (ultima > 1) sh.getRange(2, 1, ultima - 1, cab.length).clearContent();
  if (!entregas.length) return;
  const linhas = entregas.map(e => [e.id, e.pedido, e.sku, e.produto, e.un, e.qtd, e.data, e.recebidoTxt, e.entreguePor || entregador[e.pedido] || '', e.obsTxt]);
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
    escrever_(ss, CTRL.destCatalogo, dados.catalogo, 6);
    escrever_(ss, CTRL.destItens, dados.itens, 11);
    if (obras.ativo) {
      escreverObras_(ss, obras.linhas); escreverEtapas_(ss, obras.etapas);
      PropertiesService.getScriptProperties().setProperty('HASH_OBRAS', obras.assinatura);
    }
    const obrasParaVinculo = obras.ativo ? obras.linhas :
      linhas_(ss, CTRL.obras, CTRL.linhaObras, 4).filter(r => texto_(r[0]));
    const ident = identificarObras_(ss, dados, obrasParaVinculo);
    // Coluna O de _IMPORT_PEDIDOS: obra identificada pelo script (código ou rua + casa).
    escrever_(ss, CTRL.destPedidos, dados.pedidos.map((p, i) => p.concat([ident.porLinha[i], p.referencia || ''])), 16);
    const cabP = ss.getSheetByName(CTRL.destPedidos).getRange('P1');
    if (texto_(cabP.getValue()) !== 'PONTO_REFERENCIA') cabP.setValue('PONTO_REFERENCIA');
    escreverCancelamentos_(ss, dados.cancelamentos);
    escreverEntregas_(ss, dados.entregas, dados.pedidos);
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
    Logger.log(resumo);
    return resumo;
  } finally { lock.releaseLock(); }
}

function sincronizarComAviso() {
  const ui = SpreadsheetApp.getUi();
  try { ui.alert('Sincronização concluída', sincronizarFormulario(), ui.ButtonSet.OK); }
  catch (e) { ui.alert('Sincronização não concluída', e.message, ui.ButtonSet.OK); }
}

/** Acionador de 1 em 1 minuto. Verificação rápida: só faz a cópia completa quando o formulário
 *  registrou algo novo (pedido, revisão ou entrega) ou a cada 15 minutos (obras e catálogo).
 *  Assim fica quase em tempo real sem estourar o limite diário de execução do Google. */
/** Marca rápida do formulário: nº de linhas de PEDIDOS, ITENS_PEDIDO, ENTREGAS e AUDITORIA. */
function marcaOrigem_(conf) {
  const origem = SpreadsheetApp.openById(texto_(conf.getRange(CTRL.celulaOrigemId).getValue()));
  const ult = nome => { const sh = origem.getSheetByName(nome); return sh ? sh.getLastRow() : 0; };
  return [ult(texto_(conf.getRange(CTRL.celulaAbaPedidos).getValue())), ult(texto_(conf.getRange(CTRL.celulaAbaItens).getValue())),
          ult(CTRL.abaEntregasOrigem), ult('AUDITORIA')].join('|');
}

/** Chamado pelo painel antes de devolver os dados: se o formulário tem pedido, revisão ou entrega
 *  nova, sincroniza na hora. Assim o painel mostra a novidade sem esperar o acionador de 1 minuto. */
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

function sincronizacaoAutomatica() {
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

function instalarSincronizacao() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['sincronizarFormulario', 'sincronizacaoAutomatica', 'aoEditarObras'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizacaoAutomatica').timeBased().everyMinutes(CTRL.intervaloMinutos).create();
  // Na hora em que alguém muda fase, status ou etapa na planilha "Controle e Planejamento de Obras".
  const idObras = texto_(planilhaControle_().getSheetByName(CTRL.conf).getRange(CTRL.celulaObrasId).getValue());
  if (idObras) ScriptApp.newTrigger('aoEditarObras').forSpreadsheet(idObras).onEdit().create();
  PropertiesService.getScriptProperties().deleteProperty('MARCA_ORIGEM');
  const r = sincronizacaoAutomatica();
  SpreadsheetApp.getUi().alert('Sincronização automática ativada: verifica o formulário a cada minuto e copia assim que houver pedido ou entrega nova. ' +
    'Mudanças de fase, status e etapas na planilha de obras entram na hora.\n\nPrimeira execução: ' + r);
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
  try { saida = c.acao === 'entrar' ? entrarPainel(c.usuario, c.pin)
    : c.acao === 'verificarPrimeiro' ? verificarPrimeiroAcesso(c.usuario, c.codigo)
    : c.acao === 'primeiro' ? primeiroAcesso(c.usuario, c.codigo, c.pin)
    : c.acao === 'trocarPin' ? trocarPin(c.token, c.pinAtual, c.pinNovo)
    : c.acao === 'sair' ? sairAparelho(c.token)
    : c.acao === 'usuarios' ? listarUsuarios(c.token)
    : c.acao === 'salvarUsuario' ? salvarUsuario(c.token, c.dados)
    : c.acao === 'resetarPin' ? resetarPin(c.token, c.usuario)
    : c.acao === 'desconectar' ? desconectarAparelhos(c.token, c.usuario)
    : c.acao === 'entregas' ? getEntregasPedido(c.token, c.pedido)
    : c.acao === 'salvarOrcamento' ? salvarOrcamento(c.token, c.obra, c.valor, c.data, c.obs)
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

/* ===================== ACESSO INDIVIDUAL: USUÁRIO + PIN E CHAVE DO APARELHO =====================
 * Cadastro na aba USUARIOS da planilha do formulário: A EMAIL · B NOME · C USUARIO · D EMPRESA · E PERFIL ·
 * F FORMULARIO · G PAINEL · H ATIVO · I OBSERVACAO. O PIN nunca vai para a planilha nem para o navegador:
 * fica nas propriedades do script ("PIN_<usuario>" = "sal:hash"). Cada aparelho conectado tem uma chave
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
const publico_ = u => ({nome: u.nome || u.usuario, usuario: u.usuario, empresa: u.empresa, perfil: u.perfil, email: u.email,
  admin: ehAdmin_(u), adminZayda: ehAdminZayda_(u), dono: u.usuario === USUARIO_DONO});

function hashPin_(pin, sal) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin) + sal, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function guardarPin_(usuario, pin) {
  const sal = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('PIN_' + usuario, sal + ':' + hashPin_(pin, sal));
}
function pinConfere_(usuario, pin) {
  const v = PropertiesService.getScriptProperties().getProperty('PIN_' + usuario);
  if (!v) return false;
  const i = v.indexOf(':');
  return hashPin_(pin, v.slice(0, i)) === v.slice(i + 1);
}
const temPin_ = usuario => !!PropertiesService.getScriptProperties().getProperty('PIN_' + usuario);
const pinValido_ = p => /^\d{6}$/.test(texto_(p));

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
  const u = usuarioPorLogin_(reg.u);
  if (!podePainel_(u)) { props.deleteProperty('TOKEN_' + t); return {erro: {acesso: false, motivo: 'DESATIVADO'}}; }
  if (Date.now() - Number(reg.uso || 0) > 3600000) { reg.uso = Date.now(); props.setProperty('TOKEN_' + t, JSON.stringify(reg)); } // último uso, no máximo 1 gravação por hora
  return {u};
}

/** Login com usuário + PIN. 5 erros seguidos no mesmo usuário bloqueiam por 15 minutos. */
function entrarPainel(usuario, pin) {
  const login = texto_(usuario).toLowerCase(), cache = CacheService.getScriptCache();
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
  if (!u || !pinValido_(pin) || !temPin_(u.usuario) || !pinConfere_(u.usuario, texto_(pin))) {
    if (u && !temPin_(u.usuario) && podePainel_(u)) return {acesso: false, motivo: 'SEM_PIN'};
    return erro();
  }
  if (!podePainel_(u)) return {acesso: false, motivo: 'DESATIVADO'};
  cache.remove('PINERR_' + login);
  return {acesso: true, token: novoToken_(u.usuario)};
}

/** Primeiro acesso (passo 1): usuário válido, ainda sem PIN, com o código de acesso do painel.
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
/** Primeiro acesso (passo 2): cria o PIN de 6 dígitos e conecta o aparelho. */
function primeiroAcesso(usuario, codigo, pin) {
  const v = verificarPrimeiroAcesso(usuario, codigo);
  if (!v.acesso || !v.criarPin) return v.acesso ? {acesso: false, motivo: 'CODIGO_INCORRETO'} : v;
  if (!pinValido_(pin)) return {acesso: false, motivo: 'PIN_FORMATO'};
  const u = usuarioPorLogin_(usuario);
  guardarPin_(u.usuario, texto_(pin));
  return {acesso: true, token: novoToken_(u.usuario)};
}
function trocarPin(token, pinAtual, pinNovo) {
  const a = autenticar_(token); if (a.erro) return a.erro;
  if (a.u.usuario === USUARIO_DONO) return {ok: false, motivo: 'A recuperação pelo código mestre não tem PIN.'};
  if (!pinValido_(pinNovo)) return {ok: false, motivo: 'O PIN novo precisa ter 6 dígitos.'};
  const cache = CacheService.getScriptCache(), login = a.u.usuario;
  if (cache.get('PINBLOQ_' + login)) return {ok: false, motivo: 'Muitas tentativas. Tente de novo em 15 minutos.'};
  if (!pinConfere_(login, texto_(pinAtual))) {
    const n = Number(cache.get('PINERR_' + login) || 0) + 1;
    if (n >= 5) { cache.put('PINBLOQ_' + login, '1', 900); cache.remove('PINERR_' + login); return {ok: false, motivo: 'Muitas tentativas. Tente de novo em 15 minutos.'}; }
    cache.put('PINERR_' + login, String(n), 900);
    return {ok: false, motivo: 'O PIN atual está incorreto.'};
  }
  cache.remove('PINERR_' + login);
  guardarPin_(login, texto_(pinNovo));
  return {ok: true};
}
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
  const aparelhos = contarTokens_();
  return {ok: true, eu: a.u.email, dono: DONO_SISTEMA, usuarios: usuariosPainel_().map(u => ({email: u.email, nome: u.nome, usuario: u.usuario,
    empresa: u.empresa, perfil: u.perfil, formulario: u.formulario, painel: u.painel, ativo: u.ativo, obs: u.obs,
    temPin: temPin_(u.usuario), aparelhos: aparelhos[u.usuario] || 0}))};
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
    // Login trocado: o PIN e os aparelhos acompanham.
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
/** Apaga o PIN e desconecta todos os aparelhos: a pessoa faz o primeiro acesso de novo. */
function resetarPin(token, usuario) {
  const a = adminZayda_(token); if (a.erro) return a.erro;
  const g = alvoGestao_(a, usuario); if (g.erro) return g.erro;
  PropertiesService.getScriptProperties().deleteProperty('PIN_' + g.alvo.usuario);
  const n = apagarTokens_(g.alvo.usuario);
  auditarUsuario_(a.u, g.alvo.email, 'PIN do usuário ' + g.alvo.usuario + ' apagado; ' + n + ' aparelho(s) desconectado(s)');
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
  const pedidos = {};
  linhas_(ss, CTRL.destPedidos, 2, 16).filter(r => texto_(r[0])).forEach(r => {
    const id = texto_(r[0]), rev = Number(r[1]) || 0;
    if (pedidos[id] && pedidos[id].rev > rev) return;
    pedidos[id] = {rev, criado: dataHora_(r[2]), atualizado: dataHora_(r[3]), email: texto_(r[4]), atendente: texto_(r[5]),
      entregador: texto_(r[6]), endereco: texto_(r[7]), casa: texto_(r[8]), obs: texto_(r[9]), status: texto_(r[12]),
      referencia: texto_(r[15])};
  });
  // Entregas de cada pedido, com quem entregou e quem recebeu (_IMPORT_ENTREGAS), da mais antiga para a mais nova.
  const shEnt = ss.getSheetByName(CTRL.destEntregas);
  if (shEnt && shEnt.getLastRow() > 1) shEnt.getRange(2, 1, shEnt.getLastRow() - 1, 10).getValues().forEach(r => {
    const id = texto_(r[1]);
    if (!pedidos[id] || !texto_(r[0])) return;
    (pedidos[id].entregas = pedidos[id].entregas || []).push({id: texto_(r[0]), sku: texto_(r[2]), material: texto_(r[3]), un: texto_(r[4]),
      qtd: num_(r[5]) || 0, data: dataIso_(r[6]), recebidoPor: texto_(r[7]), entreguePor: texto_(r[8]), obs: texto_(r[9])});
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
  const conf = ss.getSheetByName(CTRL.conf);
  return {
    acesso: true, pedidos, nomesEtapas, orcamentos,
    usuarios: usuariosPainel_().reduce((m, u) => { m[u.email] = {nome: u.nome, empresa: u.empresa}; return m; }, {}), eu: publico_(ac.u),
    geradoEm: Utilities.formatDate(new Date(), CTRL.fuso, "yyyy-MM-dd'T'HH:mm:ss"),
    hoje: Utilities.formatDate(new Date(), CTRL.fuso, 'yyyy-MM-dd'),
    modo: texto_(conf.getRange(CTRL.celulaModo).getValue()) || 'PRODUÇÃO',
    sincronizadoEm: texto_(conf.getRange(CTRL.celulaUltimaSinc).getValue()).split(' · ')[0],
    planilhaUrl: ss.getUrl(), versaoPainel: versaoPainel_(), obras, itens, lanc
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
