// =====================================================================
//  COMANDA CENTRAL — Google Apps Script (backend da planilha COMANDA)
//  VERSÃO 4 — 07/11/2026 (estoque ligado à rotina Blend)
//   4. Aba "Movimentos": contagens e correções de estoque lançadas no app.
//      Só ACRESCENTA linhas (nunca apaga nem altera uma linha já gravada).
//   5. Aba "Estoque App": saldo calculado pela rotina (vendas x ficha técnica).
//      O app só LÊ; nenhum salvamento do app escreve nela.
//  (versão 3 — 31/10/2026)
//  Correções:
//   1. Fim dos fechamentos duplicados (merge por ID + detecção de
//      duplicata por Data/Total, com trava contra salvamentos simultâneos)
//   2. Observações de retirada / diferença de troco não se perdem mais
//      (qualquer campo novo que o app enviar ganha coluna própria e
//      campo vazio vindo de aparelho desatualizado não apaga o que já existe)
//   3. Aba "Retiradas" gerada automaticamente a partir dos fechamentos
//  Como atualizar: apague tudo no editor, cole este código, salve e
//  Implantar → Gerenciar implantações → lápis → Nova versão → Implantar.
//  Depois rode UMA VEZ a função limparDuplicados() no editor.
// =====================================================================

const SPREADSHEET_ID   = '1ttrBuvzgpcVloEt9hcHZsB7rLDwCKxc4ogsFB5Yv7tY';
const SHEET_NAME_FECH  = 'Fechamentos';
const SHEET_NAME_RET   = 'Retiradas';
const SHEET_NAME_PROD  = 'Produtos';
const SHEET_NAME_INSUM = 'Insumos';
const SHEET_NAME_FORN  = 'Fornecedores';
const SHEET_NAME_CLIEN = 'Clientes';

const SHEET_NAME_MOV   = 'Movimentos';
const SHEET_NAME_EST   = 'Estoque App';
const MOV_COLS = ['id', 'em', 'tipo', 'por', 'itens', 'motivo', 'obs', 'aparelho'];

const MIN_SAFE_ROWS = 1;

const FECH_COLS = [
  { key: 'id',                  label: 'ID' },
  { key: 'data',                label: 'Data' },
  { key: 'caixa',               label: 'Caixa' },
  { key: 'conf.pg_cartao',      label: 'Cartão' },
  { key: 'conf.pg_cartaocpf',   label: 'Cartão CPF' },
  { key: 'conf.pg_online',      label: 'Pago Online' },
  { key: 'conf.pg_voucher',     label: 'Online/Voucher' },
  { key: 'conf.pg_onlinepix',   label: 'Online via PIX' },
  { key: 'conf.pg_dinheiro',    label: 'Dinheiro' },
  { key: 'conf.pg_motoboy',     label: 'Motoboy' },
  { key: 'conf.pg_colaborador', label: 'Colaborador' },
  { key: 'totalVendas',         label: 'Total de Vendas' },
  { key: 'trocoInicio',         label: 'Troco Início' },
  { key: 'trocoFinal',          label: 'Troco Final' },
  { key: 'motoboyRepasse',      label: 'Repasse Frete Motoboy' },
  { key: 'motoboyPix',          label: 'Motoboy PIX Conta' },
  { key: 'obs',                 label: 'Observações' },
];
const BASE_KEYS   = FECH_COLS.map(c => c.key);
const NUM_KEYS    = ['totalVendas', 'trocoInicio', 'trocoFinal', 'motoboyRepasse', 'motoboyPix'];
const PAY_KEYS    = ['pg_cartao','pg_cartaocpf','pg_online','pg_voucher','pg_onlinepix','pg_dinheiro','pg_motoboy','pg_colaborador'];
const COL_ALERTA  = 'Alerta';            // coluna de controle gerada pelo script
const COL_ATUALIZ = 'Atualizado em';     // auditoria

// ---------------------------------------------------------------------
//  HTTP
// ---------------------------------------------------------------------
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'ping';
    if (action === 'ping') return jsonResponse({ ok: true, data: { ok: true, versao: 4 } });
    if (action === 'load') return jsonResponse({ ok: true, data: loadAll() });
    if (action === 'save') {
      const d = e.parameter.d;
      if (!d) return jsonResponse({ ok: false, error: 'sem dados' });
      const payload = JSON.parse(decodeURIComponent(d));
      return jsonResponse({ ok: true, data: saveAll(payload.s || payload) });
    }
    return jsonResponse({ ok: false, error: 'ação desconhecida' });
  } catch (err) {
    logErro('doGet', err);
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) return jsonResponse({ ok: false, error: 'sem body' });
    const payload = JSON.parse(e.postData.contents);
    return jsonResponse({ ok: true, data: saveAll(payload.s || payload) });
  } catch (err) {
    logErro('doPost', err);
    return jsonResponse({ ok: false, error: err.message });
  }
}

function getSheet() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function loadAll() {
  const ss = getSheet();
  return {
    fechamentos : loadFechamentos(ss).map(stripControl),
    produtos    : sheetToJson(ss, SHEET_NAME_PROD),
    insumos     : sheetToJson(ss, SHEET_NAME_INSUM),
    fornecedores: sheetToJson(ss, SHEET_NAME_FORN),
    clientes    : sheetToJson(ss, SHEET_NAME_CLIEN),
    movimentos  : sheetToJson(ss, SHEET_NAME_MOV),
    estoque     : sheetToJson(ss, SHEET_NAME_EST).filter(r => r && r.id),
  };
}

// Trava: dois aparelhos salvando ao mesmo tempo era uma das causas de
// perda de dados (o último a salvar sobrescrevia tudo).
function saveAll(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSheet();
    const saved = [], bloqueados = [];
    let resumoFech = null;

    if (payload.fechamentos !== undefined) {
      resumoFech = saveFechamentos(ss, payload.fechamentos, payload.fechamentosExcluidos || []);
      resumoFech ? saved.push('fechamentos') : bloqueados.push('fechamentos');
    }

    const map = { produtos: SHEET_NAME_PROD, insumos: SHEET_NAME_INSUM, fornecedores: SHEET_NAME_FORN, clientes: SHEET_NAME_CLIEN };
    Object.keys(map).forEach(k => {
      if (payload[k] !== undefined) {
        const ok = jsonToSheet(ss, map[k], payload[k]);
        ok ? saved.push(k) : bloqueados.push(k);
      }
    });

    if (Array.isArray(payload.movimentos) && payload.movimentos.length) {
      const n = acrescentarMovimentos(ss, payload.movimentos);
      saved.push('movimentos' + (n ? ' +' + n : ''));
    }

    if (bloqueados.length) logErro('saveAll-bloqueio', new Error('Payload vazio bloqueado para: ' + bloqueados.join(', ')));
    SpreadsheetApp.flush();
    return { saved, bloqueados, fechamentos: resumoFech };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
//  Utilitários
// ---------------------------------------------------------------------
function getVal(obj, path) {
  return path.split('.').reduce((o, k) => (o && o[k] !== undefined ? o[k] : ''), obj);
}
function setVal(obj, path, v) {
  const parts = path.split('.');
  let o = obj;
  for (let i = 0; i < parts.length - 1; i++) { if (!o[parts[i]] || typeof o[parts[i]] !== 'object') o[parts[i]] = {}; o = o[parts[i]]; }
  o[parts[parts.length - 1]] = v;
}
function isVazio(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string') return v.trim() === '';
  if (Array.isArray(v)) return v.length === 0;
  if (typeof v === 'object') return Object.keys(v).length === 0;
  return false;
}
function num(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  let s = String(v).replace(/R\$\s*/g, '').trim();
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function normData(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}
function centavos(v) { return Math.round(num(v) * 100); }

// Impressão digital do fechamento: mesma data + mesmo total + mesmas formas de pagamento
function fingerprint(f) {
  const conf = f.conf || {};
  return [normData(f.data), centavos(f.totalVendas)].concat(PAY_KEYS.map(k => centavos(conf[k]))).join('|');
}

function stripControl(f) {
  const o = JSON.parse(JSON.stringify(f));
  delete o[COL_ALERTA]; delete o[COL_ATUALIZ];
  return o;
}

// Junta o registro que chegou com o que já está salvo.
// Regra de ouro: campo VAZIO vindo do app NUNCA apaga campo preenchido.
// (Aparelho com versão antiga carregada mandava obs/retiradas vazias e apagava.)
function mergeRegistro(antigo, novo) {
  const out = JSON.parse(JSON.stringify(antigo || {}));
  Object.keys(novo || {}).forEach(k => {
    const v = novo[k];
    if (k === 'conf' && v && typeof v === 'object' && !Array.isArray(v)) {
      out.conf = out.conf || {};
      Object.keys(v).forEach(ck => { if (!isVazio(v[ck])) out.conf[ck] = v[ck]; else if (out.conf[ck] === undefined) out.conf[ck] = v[ck]; });
      return;
    }
    if (isVazio(v) && !isVazio(out[k])) return; // protege dado existente
    out[k] = v;
  });
  return out;
}

// ---------------------------------------------------------------------
//  FECHAMENTOS — merge por ID (não apaga mais o que não veio no payload)
// ---------------------------------------------------------------------
function saveFechamentos(ss, arr, excluidos) {
  let sh = ss.getSheetByName(SHEET_NAME_FECH);
  if (!sh) sh = ss.insertSheet(SHEET_NAME_FECH);

  if ((!arr || arr.length === 0) && (!excluidos || excluidos.length === 0) && sh.getLastRow() > MIN_SAFE_ROWS) return null;

  const existentes = loadFechamentos(ss);
  const porId = {};
  const porFp = {};
  const ordem = [];
  existentes.forEach(f => {
    const id = String(f.id || '');
    if (!id) return;
    porId[id] = f; ordem.push(id);
    const fp = fingerprint(f); if (!porFp[fp]) porFp[fp] = id;
  });

  const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const resumo = { novos: [], atualizados: [], duplicadosIgnorados: [], excluidos: [] };

  (arr || []).forEach(raw => {
    if (!raw || typeof raw !== 'object') return;
    const f = JSON.parse(JSON.stringify(raw));
    delete f[COL_ALERTA]; delete f[COL_ATUALIZ];
    f.data = normData(f.data);
    let id = String(f.id || '');

    if (id && porId[id]) {
      const antes = JSON.stringify(stripControl(porId[id]));
      const depois = mergeRegistro(porId[id], f);
      if (JSON.stringify(stripControl(depois)) !== antes) { depois[COL_ATUALIZ] = agora; resumo.atualizados.push(id); }
      porId[id] = depois;
      return;
    }

    // ID novo: é a mesma venda já salva com outro ID? (clique duplo, reenvio, aparelho offline)
    const fp = fingerprint(f);
    if (porFp[fp]) {
      const alvo = porFp[fp];
      porId[alvo] = mergeRegistro(porId[alvo], Object.assign({}, f, { id: alvo }));
      resumo.duplicadosIgnorados.push((id || '(sem id)') + '→' + alvo);
      return;
    }

    if (!id) id = Utilities.getUuid().slice(0, 7);
    f.id = id;
    f[COL_ATUALIZ] = agora;
    porId[id] = f; ordem.push(id); porFp[fp] = id;
    resumo.novos.push(id);
  });

  (excluidos || []).forEach(id => {
    id = String(id);
    if (porId[id]) { delete porId[id]; resumo.excluidos.push(id); }
  });

  const lista = ordem.filter(id => porId[id]).map(id => porId[id]);

  // Alerta para datas repetidas com valores diferentes (não apaga — só sinaliza)
  const contaData = {};
  lista.forEach(f => { const d = normData(f.data); contaData[d] = (contaData[d] || 0) + 1; });
  lista.forEach(f => { f[COL_ALERTA] = contaData[normData(f.data)] > 1 ? 'DATA REPETIDA — conferir' : ''; });

  lista.sort((a, b) => normData(a.data) < normData(b.data) ? -1 : normData(a.data) > normData(b.data) ? 1 : 0);

  escreverFechamentos(sh, lista);
  escreverRetiradas(ss, lista);

  if (resumo.duplicadosIgnorados.length) logErro('duplicata-bloqueada', new Error(resumo.duplicadosIgnorados.join(', ')));
  return resumo;
}

// Escreve a aba inteira a partir do estado JÁ MESCLADO no servidor.
// Colunas extras (qualquer campo que o app enviar além das 17 originais)
// são criadas automaticamente — é aqui que as observações de retirada e de
// troco deixavam de ser salvas na versão anterior.
function escreverFechamentos(sh, lista) {
  const extras = [];
  lista.forEach(f => {
    Object.keys(f).forEach(k => {
      if (k === 'conf') {
        Object.keys(f.conf || {}).forEach(ck => { const p = 'conf.' + ck; if (BASE_KEYS.indexOf(p) < 0 && extras.indexOf(p) < 0) extras.push(p); });
        return;
      }
      if (BASE_KEYS.indexOf(k) < 0 && extras.indexOf(k) < 0 && k !== COL_ALERTA && k !== COL_ATUALIZ) extras.push(k);
    });
  });
  // só cria coluna extra se pelo menos um fechamento tiver algo nela
  const extrasUteis = extras.filter(k => lista.some(f => !isVazio(getVal(f, k))));
  extras.length = 0; extrasUteis.forEach(k => extras.push(k));
  // mantém a ordem das colunas extras que já existiam na planilha
  const headerAtual = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  const extrasOrdenadas = headerAtual.filter(h => extras.indexOf(h) >= 0).concat(extras.filter(k => headerAtual.indexOf(k) < 0));

  const headers = FECH_COLS.map(c => c.label).concat(extrasOrdenadas, [COL_ALERTA, COL_ATUALIZ]);
  const keys    = BASE_KEYS.concat(extrasOrdenadas, [COL_ALERTA, COL_ATUALIZ]);

  const rows = lista.map(f => keys.map(k => {
    const v = (k === COL_ALERTA || k === COL_ATUALIZ) ? f[k] : getVal(f, k);
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') return isVazio(v) ? '' : JSON.stringify(v);
    return v;
  }));

  // monta tudo ANTES de limpar: se der erro acima, a planilha fica intacta
  sh.clearContents();
  if (rows.length) sh.getRange(2, 2, rows.length, 1).setNumberFormat('@'); // Data como texto AAAA-MM-DD
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const hr = sh.getRange(1, 1, 1, headers.length);
  hr.setBackground('#2c2218'); hr.setFontColor('#e0a83e'); hr.setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2, 4, rows.length, 12).setNumberFormat('R$ #,##0.00'); // D..O
  }
  sh.setFrozenRows(1);
}

// Aba "Retiradas": uma linha por retirada lançada no fechamento
function escreverRetiradas(ss, lista) {
  const linhas = [];
  lista.forEach(f => {
    Object.keys(f).forEach(k => {
      if (!/sa[ií]da|retirad/i.test(k)) return;
      let v = f[k];
      if (typeof v === 'string' && /^\s*\[/.test(v)) { try { v = JSON.parse(v); } catch (e) {} }
      if (!Array.isArray(v)) return;
      v.forEach(r => {
        if (r === null || r === undefined) return;
        if (typeof r !== 'object') { linhas.push([f.id, normData(f.data), num(r), '', f.caixa || '', k]); return; }
        const valor = r.valor !== undefined ? r.valor : (r.value !== undefined ? r.value : (r.v !== undefined ? r.v : r.total));
        const obs = r.obs || r.observacao || r.observação || r.desc || r.descricao || r.motivo || r.nome || '';
        linhas.push([f.id, normData(f.data), num(valor), obs, r.operador || f.caixa || '', JSON.stringify(r)]);
      });
    });
  });
  let sh = ss.getSheetByName(SHEET_NAME_RET);
  if (!sh) sh = ss.insertSheet(SHEET_NAME_RET);
  sh.clearContents();
  const head = ['ID Fechamento', 'Data', 'Valor', 'Observação', 'Operador', 'Detalhe (origem)'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (linhas.length) {
    sh.getRange(2, 1, linhas.length, head.length).setValues(linhas);
    sh.getRange(2, 3, linhas.length, 1).setNumberFormat('R$ #,##0.00');
  }
}

function loadFechamentos(ss) {
  const sh = ss.getSheetByName(SHEET_NAME_FECH);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);
  const labelToKey = {};
  FECH_COLS.forEach(c => labelToKey[c.label] = c.key);

  return data.slice(1).filter(r => String(r[0]).trim() !== '').map(row => {
    const obj = { conf: {}, saidas: [], vendas: [] };
    headers.forEach((h, i) => {
      if (!h) return;
      const key = labelToKey[h] || h;   // colunas extras: o cabeçalho é a própria chave
      let v = row[i];
      if (key.indexOf('conf.') === 0 && BASE_KEYS.indexOf(key) >= 0) { obj.conf[key.slice(5)] = num(v); return; }
      if (NUM_KEYS.indexOf(key) >= 0) { obj[key] = num(v); return; }
      if (key === 'data') { obj.data = normData(v); return; }
      if (typeof v === 'string' && /^\s*[\[{]/.test(v)) { try { v = JSON.parse(v); } catch (e) {} }
      if (v === '' && Array.isArray(obj[key])) return; // mantém [] padrão
      setVal(obj, key, v);
    });
    if (obj.id !== undefined) obj.id = String(obj.id);
    return obj;
  });
}

// ---------------------------------------------------------------------
//  Abas simples (produtos, insumos, fornecedores, clientes)
// ---------------------------------------------------------------------
function sheetToJson(ss, name) {
  const sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => { const obj = {}; headers.forEach((h, i) => { try { obj[h] = JSON.parse(row[i]); } catch (e) { obj[h] = row[i]; } }); return obj; });
}

function jsonToSheet(ss, name, arr) {
  let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
  if ((!arr || arr.length === 0) && sh.getLastRow() > MIN_SAFE_ROWS) return false;
  if (!arr || arr.length === 0) { sh.clearContents(); return true; }
  const keys = []; arr.forEach(obj => Object.keys(obj).forEach(k => { if (keys.indexOf(k) < 0) keys.push(k); }));
  const rows = [keys].concat(arr.map(obj => keys.map(k => { const v = obj[k]; return (typeof v === 'object' && v !== null) ? JSON.stringify(v) : (v === undefined ? '' : v); })));
  sh.clearContents();
  sh.getRange(1, 1, rows.length, keys.length).setValues(rows);
  sh.autoResizeColumns(1, keys.length);
  return true;
}

// ---------------------------------------------------------------------
//  RODAR UMA VEZ: remove os duplicados que já estão na planilha
//  (faz backup antes). Duplicado = mesma data + mesmo total + mesmas
//  formas de pagamento. Mantém o primeiro, junta as observações.
// ---------------------------------------------------------------------
function limparDuplicados() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    backupDiario();
    const ss = getSheet();
    const lista = loadFechamentos(ss);
    const porFp = {}, final = [], removidos = [];
    lista.forEach(f => {
      const fp = fingerprint(f);
      if (porFp[fp]) { const i = porFp[fp] - 1; final[i] = mergeRegistro(final[i], Object.assign({}, f, { id: final[i].id })); removidos.push(f.id + ' (' + f.data + ')'); return; }
      final.push(f); porFp[fp] = final.length;
    });
    const contaData = {};
    final.forEach(f => { contaData[f.data] = (contaData[f.data] || 0) + 1; });
    final.forEach(f => { f[COL_ALERTA] = contaData[f.data] > 1 ? 'DATA REPETIDA — conferir' : ''; });
    final.sort((a, b) => a.data < b.data ? -1 : a.data > b.data ? 1 : 0);
    escreverFechamentos(ss.getSheetByName(SHEET_NAME_FECH), final);
    escreverRetiradas(ss, final);
    const msg = removidos.length ? 'Removidos ' + removidos.length + ' duplicados: ' + removidos.join(', ') : 'Nenhum duplicado encontrado';
    logErro('limparDuplicados', new Error(msg));
    Logger.log(msg);
    return msg;
  } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------------
//  BACKUP AUTOMÁTICO DIÁRIO — rode criarTriggerBackup() uma vez
// ---------------------------------------------------------------------
function backupDiario() {
  const arquivo = DriveApp.getFileById(SPREADSHEET_ID);
  const pais = arquivo.getParents();
  const pasta = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  const nome = 'Backup Comanda Central - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  arquivo.makeCopy(nome, pasta);
  const todos = pasta.getFiles(); const backups = [];
  while (todos.hasNext()) { const f = todos.next(); if (f.getName().indexOf('Backup Comanda Central - ') === 0) backups.push(f); }
  backups.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  backups.slice(14).forEach(f => f.setTrashed(true));
}

function criarTriggerBackup() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'backupDiario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupDiario').timeBased().everyDays(1).atHour(4).create();
}

// ---------------------------------------------------------------------
//  LOG — aba "_Log"
// ---------------------------------------------------------------------
function logErro(origem, err) {
  try {
    const ss = getSheet();
    let sh = ss.getSheetByName('_Log');
    if (!sh) { sh = ss.insertSheet('_Log'); sh.appendRow(['Data/Hora', 'Origem', 'Mensagem']); }
    sh.appendRow([new Date(), origem, err && err.message ? err.message : String(err)]);
  } catch (e) { /* não quebra a resposta principal */ }
}

// ---------------------------------------------------------------------
//  MOVIMENTOS DE ESTOQUE (v4) — só acrescenta, por ID
//  Contagem nova / correção lançadas pela gerente no app. A rotina Blend
//  lê esta aba e grava o saldo calculado na aba "Estoque App".
// ---------------------------------------------------------------------
function acrescentarMovimentos(ss, arr) {
  let sh = ss.getSheetByName(SHEET_NAME_MOV);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME_MOV); sh.getRange(1, 1, 1, MOV_COLS.length).setValues([MOV_COLS]).setFontWeight('bold'); sh.setFrozenRows(1); }
  const n = sh.getLastRow();
  const ids = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(r => { ids[String(r[0])] = 1; });
  const novas = [];
  arr.forEach(m => {
    if (!m || typeof m !== 'object' || !m.id || ids[String(m.id)]) return;
    ids[String(m.id)] = 1;
    novas.push(MOV_COLS.map(k => { const v = m[k]; return (v && typeof v === 'object') ? JSON.stringify(v) : (v === undefined || v === null ? '' : v); }));
  });
  if (novas.length) {
    sh.getRange(n + 1, 1, novas.length, 1).setNumberFormat('@');
    sh.getRange(n + 1, 2, novas.length, 1).setNumberFormat('@');
    sh.getRange(n + 1, 1, novas.length, MOV_COLS.length).setValues(novas);
  }
  return novas.length;
}
