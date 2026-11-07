// ================================================================
// MEU FINANCEIRO v1.0.0 — DIAGNOSTICO.gs
// Rode DIAG_COMPLETO() e me envie o log se algo não funcionar.
// ================================================================

function DIAG_COMPLETO() {
  var L = [];
  function ok(t) { L.push('✅ ' + t); }
  function er(t) { L.push('❌ ' + t); }
  function av(t) { L.push('⚠️ ' + t); }
  L.push('===== DIAGNÓSTICO ' + APP_NOME + ' v' + VERSAO + ' — ' + _agora() + ' =====');

  // 0. Conta (v1.1.0)
  try {
    var em = _emailAtual();
    (em ? ok : er)('Conta identificada: ' + (em || 'NENHUMA'));
    var ssx = _ss(); var dn = '';
    try { dn = DriveApp.getFileById(ssx.getId()).getOwner().getEmail(); } catch (e) {}
    (String(dn).toLowerCase() === em ? ok : av)('Dono da planilha: ' + (dn || 'não identificado'));
    (_lerConfig().DONO_EMAIL === em ? ok : er)('Carimbo DONO_EMAIL: ' + _lerConfig().DONO_EMAIL);
    ok('Tema salvo: ' + (_lerConfig().TEMA_COR || 'violeta') + ' / ' + (_lerConfig().TEMA_MODO || 'dark'));
  } catch (e) { er('Conta: ' + e.message); }

  // 1. Planilha e abas
  try {
    var ss = _garantirBase(); ok('Planilha: ' + ss.getUrl());
    Object.keys(ABAS).forEach(function (n) {
      var sh = ss.getSheetByName(n);
      if (!sh) return er('Aba ausente: ' + n);
      var cab = sh.getRange(1, 1, 1, ABAS[n].length).getValues()[0].join('|');
      if (cab !== ABAS[n].join('|')) er('Cabeçalho diferente em ' + n + ': ' + cab);
      else ok(n + ': ' + Math.max(0, sh.getLastRow() - 1) + ' linha(s)');
    });
  } catch (e) { er('Planilha: ' + e.message); }

  // 2. Config
  try {
    var cfg = _lerConfig();
    Object.keys(CONFIG_PADRAO).forEach(function (k) { if (cfg[k] === undefined) av('Config sem chave ' + k); });
    ok('Config lida (' + Object.keys(cfg).length + ' chaves) · IA: ' + (cfg.GEMINI_API_KEY ? 'chave cadastrada' : 'sem chave (usa OCR)'));
  } catch (e) { er('Config: ' + e.message); }

  // 3. Painel do mês
  try {
    var t0 = new Date().getTime();
    var p = _painel();
    ok('Painel ' + p.mes + ' em ' + (new Date().getTime() - t0) + ' ms · receita ' + _brl(p.kpis.receita) + ' · gastos ' + _brl(p.kpis.gastos) + ' · score ' + p.kpis.score);
    ok('Cartões ' + p.cartoes.length + ' · parcelamentos ' + p.parcelamentos.length + ' · financiamentos ' + p.financiamentos.length + ' · lembretes ' + p.lembretes.length);
    var tam = JSON.stringify(p).length; (tam > 4000000 ? av : ok)('Tamanho do retorno: ' + Math.round(tam / 1024) + ' KB');
  } catch (e) { er('Painel: ' + e.message + ' ' + (e.stack || '')); }

  // 4. Competência de cartão
  try {
    var c = { DIA_FECHAMENTO: 25, DIA_VENCIMENTO: 5 };
    var a = _competencia('2026-11-10', c), b = _competencia('2026-11-26', c);
    (a === '2026-12' && b === '2027-01' ? ok : er)('Competência cartão (fecha 25/vence 5): 10/11→' + a + ' · 26/11→' + b);
  } catch (e) { er('Competência: ' + e.message); }

  // 5. API (todas as funções existem)
  ['apiPainel', 'apiTodosLancamentos', 'apiSalvarLancamento', 'apiExcluirLancamento', 'apiQuitarGrupo', 'apiSalvarCartao', 'apiExcluirCartao', 'apiSalvarReceita', 'apiExcluirReceita', 'apiCalcProlabore', 'apiSalvarFinanciamento', 'apiPagarParcelaFin', 'apiExcluirFinanciamento', 'apiSalvarContaFixa', 'apiExcluirContaFixa', 'apiSalvarLembrete', 'apiConcluirLembrete', 'apiEnviarLembreteEmail', 'apiExcluirLembrete', 'apiSalvarMeta', 'apiAportarMeta', 'apiExcluirMeta', 'apiLerNota', 'apiMarcarNotificacoes', 'apiLimparNotificacoes', 'apiSalvarConfig', 'apiTestarIA', 'apiExportarMes', 'apiConta']
    .forEach(function (f) { if (typeof this[f] !== 'function') er('Função ausente: ' + f); });
  ok('Funções de API verificadas');

  // 6. Gatilhos
  var tg = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  ['rotinaDiaria', 'rotinaLembretes'].forEach(function (f) { (tg.indexOf(f) > -1 ? ok : er)('Gatilho ' + f + (tg.indexOf(f) > -1 ? ' instalado' : ' AUSENTE — rode SETUP_INICIAL')); });

  // 7. Pasta das notas + OCR + IA
  try { ok('Pasta notas: ' + _pastaNotas().getUrl()); } catch (e) { er('Pasta notas: ' + e.message); }
  try {
    var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    (r.getResponseCode() === 200 ? ok : er)('Acesso à API do Drive (OCR): HTTP ' + r.getResponseCode());
  } catch (e) { er('API Drive: ' + e.message); }
  try {
    var cf = _lerConfig();
    if (cf.GEMINI_API_KEY) {
      var g = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(cf.GEMINI_API_KEY), { muteHttpExceptions: true });
      (g.getResponseCode() === 200 ? ok : er)('Gemini: HTTP ' + g.getResponseCode());
    } else av('Gemini sem chave — leitura será por OCR do Drive');
  } catch (e) { er('Gemini: ' + e.message); }

  // 8. Cotas
  try { ok('Cota de e-mails restante hoje: ' + MailApp.getRemainingDailyQuota()); } catch (e) { av('Cota e-mail: ' + e.message); }

  // 9. Frontend
  try {
    var html = HtmlService.createTemplateFromFile('index').getRawContent();
    ok('index.html: ' + html.length + ' caracteres');
    if (/[\u201C\u201D\u2018\u2019]/.test(html.replace(/>[^<]*</g, '><'))) av('Possíveis aspas curvas no código do index.html');
    ['</style>', '</head>', '</body>', '</html>', 'function carregar', 'function rInicio', 'function abrirScanner', 'function enviarNota'].forEach(function (m) { if (html.indexOf(m) === -1) er('index.html sem: ' + m); });
    var aS = (html.match(/<script/g) || []).length, fS = (html.match(/<\/script>/g) || []).length;
    (aS === fS ? ok : er)('Tags script: ' + aS + ' abertas / ' + fS + ' fechadas');
  } catch (e) { er('index.html: ' + e.message); }

  L.push('===== FIM =====');
  Logger.log(L.join('\n'));
  return L.join('\n');
}

// Cria dados de exemplo para ver o painel funcionando (pode apagar depois em LIMPAR_EXEMPLOS)
function CRIAR_EXEMPLOS() {
  _garantirBase();
  var ym = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  var c1 = JSON.parse(apiSalvarCartao({ NOME: 'Nubank', BANDEIRA: 'Master', LIMITE: 8000, DIA_FECHAMENTO: 28, DIA_VENCIMENTO: 5, COR: '#8B5CF6' })).id;
  var c2 = JSON.parse(apiSalvarCartao({ NOME: 'Itaú', BANDEIRA: 'Visa', LIMITE: 12000, DIA_FECHAMENTO: 20, DIA_VENCIMENTO: 28, COR: '#F97316' })).id;
  _salvarReceita({ DATA: ym + '-05', TIPO: 'SALARIO', DESCRICAO: 'Salário', BRUTO: 9000, DESCONTOS: 1800, RECORRENTE: 'SIM', OBS: 'EXEMPLO' });
  _salvarReceita({ DATA: ym + '-10', TIPO: 'PROLABORE', DESCRICAO: 'Pró-labore', BRUTO: 3000, DESCONTOS: 330, RECORRENTE: 'SIM', OBS: 'EXEMPLO' });
  [['Mercado do mês', 'Mercado', 820, c1, 1], ['iPhone', 'Eletrônicos', 6000, c2, 10], ['Jantar', 'Restaurante', 230, c1, 1], ['Combustível', 'Combustível', 280, '', 1], ['Passagem aérea', 'Viagem', 1800, c1, 6], ['Netflix', 'Assinaturas', 55.9, c2, 1]]
    .forEach(function (x, i) { _salvarLancamento({ DATA: ym + '-' + ('0' + (3 + i * 3)).slice(-2), DESCRICAO: x[0], CATEGORIA: x[1], VALOR: x[2], CARTAO_ID: x[3], FORMA: x[3] ? 'Crédito' : 'Pix', TOTAL_PARCELAS: x[4], OBS: 'EXEMPLO' }); });
  apiSalvarFinanciamento({ NOME: 'Carro', TIPO: 'Veículo', INSTITUICAO: 'Banco', VALOR_TOTAL: 60000, VALOR_PARCELA: 1450, TOTAL_PARCELAS: 48, PARCELAS_PAGAS: 12, DIA_VENCIMENTO: 15, INICIO: Utilities.formatDate(new Date(new Date().getFullYear() - 1, new Date().getMonth(), 15), TZ, 'yyyy-MM-dd'), OBS: 'EXEMPLO' });
  apiSalvarContaFixa({ NOME: 'Aluguel', CATEGORIA: 'Moradia', VALOR: 2200, DIA_VENCIMENTO: 10, OBS: 'EXEMPLO' });
  apiSalvarContaFixa({ NOME: 'Internet', CATEGORIA: 'Contas da casa', VALOR: 120, DIA_VENCIMENTO: 18, OBS: 'EXEMPLO' });
  _salvarLembrete({ TITULO: 'Enviar extrato e notas do mês', MENSAGEM: 'Olá! Segue o extrato do mês e as notas fiscais. Qualquer dúvida me avise. Obrigado!', DESTINATARIO: 'Contador', DATA: ym + '-05', HORA: '09:00', RECORRENCIA: 'MENSAL' });
  apiSalvarMeta({ NOME: 'Reserva de emergência', VALOR_ALVO: 30000, VALOR_ATUAL: 8500, COR: '#10B981' });
  Logger.log('✅ Exemplos criados. Abra o Web App.');
}

function LIMPAR_EXEMPLOS() {
  var n = 0;
  ['LANCAMENTOS', 'RECEITAS', 'FINANCIAMENTOS', 'CONTAS_FIXAS'].forEach(function (a) { n += _excluirOnde(a, function (r) { return String(r.OBS).indexOf('EXEMPLO') > -1; }); });
  n += _excluirOnde('CARTOES', function (r) { return ['Nubank', 'Itaú'].indexOf(r.NOME) > -1; });
  n += _excluirOnde('LEMBRETES', function (r) { return r.TITULO === 'Enviar extrato e notas do mês'; });
  n += _excluirOnde('METAS', function (r) { return r.NOME === 'Reserva de emergência'; });
  Logger.log('🧹 ' + n + ' registro(s) de exemplo removidos.');
}

function TESTAR_NOTIFICACOES_AGORA() { rotinaDiaria(); rotinaLembretes(); Logger.log('✅ Rotinas executadas. Veja o sino no app e seu e-mail.'); }

