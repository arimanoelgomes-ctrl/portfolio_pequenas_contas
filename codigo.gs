// ============================================================
// PORTFÓLIO PEQUENAS CONTAS — Apps Script (Web App + Coleta Jira)
// Versão: 2.0 | Data: 09/04/2026
// ============================================================
//
// SETUP (executar UMA vez):
//
// 1. Menu Extensões → Apps Script → cole este código e salve
// 2. Configurações → Propriedades do script → adicione:
//    JIRA_BASE_URL   → https://atendimento.betha.com.br
//    JIRA_EMAIL      → seu-email@betha.com.br
//    JIRA_API_TOKEN  → token em https://id.atlassian.com/manage/api-tokens
//    SHEET_ID        → 1MKsApbL7IPf5jAsAO9N03AxAcC3ptzYbrgNQrOr_R4s
// 3. Execute testarConexaoJira() para validar o acesso
// 4. Execute onTimeTrigger() manualmente para popular Jira_Chamados
// 5. Implantar → Nova implantação:
//      Tipo: Aplicativo da Web
//      Executar como: Eu (seu usuário)
//      Quem tem acesso: Qualquer pessoa
//    Copie a URL gerada → cole em CONFIG.APPS_SCRIPT_URL no index.html
// 6. Execute setupTrigger() para ativar atualização automática (6h)
// ============================================================

const SHEET_ID_DEFAULT = '1MKsApbL7IPf5jAsAO9N03AxAcC3ptzYbrgNQrOr_R4s';

// ────────────────────────────────────────────────────────────
// WEB APP — Serve dados da planilha como JSON/JSONP
// Dashboard chama: ?sheet=Jira_Chamados&callback=fn
// ────────────────────────────────────────────────────────────
function doGet(e) {
  const callback  = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  const sheetName = (e && e.parameter && e.parameter.sheet)    ? e.parameter.sheet    : 'Jira_Chamados';

  let payload;
  try {
    const data = readSheetData(sheetName);
    payload = JSON.stringify({ status: 'ok', sheet: sheetName, data: data });
  } catch (err) {
    payload = JSON.stringify({ status: 'error', message: err.message });
  }

  const output = callback
    ? ContentService.createTextOutput(`${callback}(${payload})`).setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);

  return output;
}

// Lê uma aba e retorna array de objetos com os cabeçalhos como chaves
function readSheetData(sheetName) {
  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  const tab     = ss.getSheetByName(sheetName);

  if (!tab) throw new Error(`Aba "${sheetName}" não encontrada`);

  const values = tab.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => h.toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove acentos
    .replace(/\s+/g, '_'));

  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const val = row[i];
        obj[h] = (val === null || val === undefined) ? '' : val.toString().trim();
      });
      return obj;
    });
}

// ────────────────────────────────────────────────────────────
// JQL — Filtro completo do portfólio de pequenas contas
// ────────────────────────────────────────────────────────────
const JQL = `category = "Projetos ativos de atendimento - Filial" AND resolution = Unresolved AND issuetype not in (Melhoria, "Melhoria (sub-tarefa)") AND "Equipe responsável" in (Suporte, Residente, Serviço) AND (Vertical not in (Saúde, Educação) AND Município in ("Abdon Batista", Agrolândia, "Anita Garibaldi", Angelina, Anchieta, "Balneário Arroio do Silva", "Balneário Barra do Sul", "Balneário Camboriú", "Balneário Piçarras", Bandeirante, "Barra Bonita", "Barra Velha", "Bela Vista do Toldo", Belmonte, "Benedito Novo", Brunópolis, Caçador, Calmon, "Campo Alegre", "Capão Alto", Chapecó, Concórdia, "Dona Emma", "Erval Velho", Ermo, "Frei Rogério", Iraceminha, Imbuia, Ipira, Ipuaçu, Itá, Itajaí, Jupiá, Lacerdópolis, "Lajeado Grande", "Leoberto Leal", "Lindóia do Sul", "Luiz Alves", Luzerna, Mafra, Massaranduba, Meleiro, Modelo, "Morro da Fumaça", "Morro Grande", Penha, Peritiba, "Pescaria Brava", Pomerode, "Praia Grande", "Rio do Sul", "Rio Fortuna", "Rio Rufino", Saltinho, "Santa Terezinha", "São Bernardino", "São Bonifácio", "São Cristovão do Sul", "São João do Oeste", "São José do Cedro", "São Martinho", "São Miguel da Boa Vista", "São Pedro de Alcântara", Tangará, "Treze de Maio", Tigrinhos, Timbó, Treviso, Videira) OR Vertical in (Saúde, Educação) AND "Equipe responsável" not in (Suporte) AND Município in ("Abdon Batista", Agrolândia, "Anita Garibaldi", Angelina, Anchieta, "Balneário Arroio do Silva", "Balneário Barra do Sul", "Balneário Camboriú", "Balneário Piçarras", Bandeirante, "Barra Bonita", "Barra Velha", "Bela Vista do Toldo", Belmonte, "Benedito Novo", Brunópolis, Caçador, Calmon, "Campo Alegre", "Capão Alto", Chapecó, Concórdia, "Dona Emma", "Erval Velho", Ermo, "Frei Rogério", Iraceminha, Imbuia, Ipira, Ipuaçu, Itá, Itajaí, Jupiá, Lacerdópolis, "Lajeado Grande", "Leoberto Leal", "Lindóia do Sul", "Luiz Alves", Luzerna, Mafra, Massaranduba, Meleiro, Modelo, "Morro da Fumaça", "Morro Grande", Penha, Peritiba, "Pescaria Brava", Pomerode, "Praia Grande", "Rio do Sul", "Rio Fortuna", "Rio Rufino", Saltinho, "Santa Terezinha", "São Bernardino", "São Bonifácio", "São Cristovão do Sul", "São João do Oeste", "São José do Cedro", "São Martinho", "São Miguel da Boa Vista", "São Pedro de Alcântara", Tangará, "Treze de Maio", Tigrinhos, Timbó, Treviso, Videira))`;

const FIELD_MUNICIPIO = 'customfield_10331'; // Município (string)
const FIELD_VERTICAL  = 'customfield_10300'; // Vertical  ({ value: "Saúde" })
const SHEET_TAB_NAME  = 'Jira_Chamados';
const PAGE_SIZE       = 100;
const SLEEP_MS        = 200;

// ────────────────────────────────────────────────────────────
// ENTRY POINT — disparado automaticamente pelo trigger (6h)
// ────────────────────────────────────────────────────────────
function onTimeTrigger() {
  const inicio = new Date();
  Logger.log(`▶ Iniciando coleta Jira: ${inicio.toLocaleString('pt-BR')}`);
  try {
    const issues = fetchJiraIssues();
    Logger.log(`  Issues coletadas: ${issues.length}`);
    const rows = aggregateByMunicipioVertical(issues);
    writeJiraChamados(rows);
    Logger.log(`✅ Concluído em ${Math.round((new Date()-inicio)/1000)}s — ${rows.length} linhas gravadas`);
  } catch (e) {
    Logger.log(`❌ ERRO: ${e.message}`);
    throw e;
  }
}

// ────────────────────────────────────────────────────────────
// AUTENTICAÇÃO — cria sessão Jira e retorna cookie JSESSIONID
// Jira Data Center usa session cookie, não Basic Auth
// ────────────────────────────────────────────────────────────
function getJiraSession(baseUrl, email, password) {
  const resp = UrlFetchApp.fetch(`${baseUrl}/rest/auth/1/session`, {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    payload: JSON.stringify({ username: email, password: password }),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error(`Jira login falhou HTTP ${code}: ${resp.getContentText().substring(0, 200)}`);
  }
  const data = JSON.parse(resp.getContentText());
  return `${data.session.name}=${data.session.value}`;
}

// ────────────────────────────────────────────────────────────
// COLETA JIRA — Jira REST API v2 com paginação
// ────────────────────────────────────────────────────────────
function fetchJiraIssues() {
  const props    = PropertiesService.getScriptProperties();
  const baseUrl  = props.getProperty('JIRA_BASE_URL') || '';
  const email    = props.getProperty('JIRA_EMAIL')    || '';
  const password = props.getProperty('JIRA_API_TOKEN') || '';

  if (!baseUrl || !email || !password) {
    throw new Error('Script Properties incompletas: configure JIRA_BASE_URL, JIRA_EMAIL e JIRA_API_TOKEN.');
  }

  const sessionCookie = getJiraSession(baseUrl, email, password);
  const headers = { 'Cookie': sessionCookie, 'Accept': 'application/json', 'Content-Type': 'application/json' };
  const fields  = ['summary', FIELD_MUNICIPIO, FIELD_VERTICAL];

  const allIssues = [];
  let startAt = 0;

  while (true) {
    const url  = `${baseUrl}/rest/api/2/search`;
    const body = JSON.stringify({ jql: JQL, fields: fields, maxResults: PAGE_SIZE, startAt: startAt });

    const resp = UrlFetchApp.fetch(url, { method: 'post', headers, payload: body, muteHttpExceptions: true });
    const code = resp.getResponseCode();

    if (code !== 200) {
      throw new Error(`Jira HTTP ${code}: ${resp.getContentText().substring(0, 300)}`);
    }

    const data = JSON.parse(resp.getContentText());
    allIssues.push(...data.issues);
    Logger.log(`  Página ${Math.floor(startAt/PAGE_SIZE)+1}: ${data.issues.length} issues (${allIssues.length}/${data.total})`);

    if (allIssues.length >= data.total || data.issues.length === 0) break;
    startAt += PAGE_SIZE;
    Utilities.sleep(SLEEP_MS);
  }

  return allIssues;
}

// ────────────────────────────────────────────────────────────
// AGREGAÇÃO — agrupa por Município × Vertical
// ────────────────────────────────────────────────────────────
function aggregateByMunicipioVertical(issues) {
  const map = {};
  const ts  = new Date().toISOString();

  issues.forEach(issue => {
    const municipio = issue.fields[FIELD_MUNICIPIO] || 'Não informado';
    const vObj      = issue.fields[FIELD_VERTICAL];
    const vertical  = (vObj && vObj.value) ? vObj.value : 'Não informado';
    const key       = `${municipio}||${vertical}`;
    map[key]        = (map[key] || 0) + 1;
  });

  return Object.entries(map).map(([key, count]) => {
    const [municipio, vertical] = key.split('||');
    return [municipio, vertical, count, ts];
  });
}

// ────────────────────────────────────────────────────────────
// GRAVAÇÃO — atualiza a aba Jira_Chamados
// ────────────────────────────────────────────────────────────
function writeJiraChamados(rows) {
  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  let tab       = ss.getSheetByName(SHEET_TAB_NAME);
  if (!tab) { tab = ss.insertSheet(SHEET_TAB_NAME); Logger.log(`  Aba "${SHEET_TAB_NAME}" criada.`); }

  // Cabeçalho
  tab.getRange(1, 1, 1, 4).setValues([['municipio','vertical','total_chamados','atualizado_em']]);

  // Limpar dados anteriores
  const last = tab.getLastRow();
  if (last > 1) tab.getRange(2, 1, last - 1, 4).clearContent();

  // Gravar
  if (rows.length > 0) tab.getRange(2, 1, rows.length, 4).setValues(rows);

  // Estilo no cabeçalho
  const h = tab.getRange(1, 1, 1, 4);
  h.setBackground('#1E3A5F');
  h.setFontColor('#FFFFFF');
  h.setFontWeight('bold');
  tab.setFrozenRows(1);

  Logger.log(`  "${SHEET_TAB_NAME}" atualizada: ${rows.length} linhas.`);
}

// ────────────────────────────────────────────────────────────
// SETUP — instala trigger automático (executar UMA vez)
// ────────────────────────────────────────────────────────────
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onTimeTrigger')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onTimeTrigger').timeBased().everyHours(6).create();
  Logger.log('✅ Trigger instalado: onTimeTrigger a cada 6 horas.');
}

// ────────────────────────────────────────────────────────────
// DIAGNÓSTICO — testa conexão com Jira (não grava nada)
// ────────────────────────────────────────────────────────────
function testarConexaoJira() {
  const props   = PropertiesService.getScriptProperties();
  const baseUrl = props.getProperty('JIRA_BASE_URL') || '';
  const email   = props.getProperty('JIRA_EMAIL')    || '';
  const token   = props.getProperty('JIRA_API_TOKEN') || '';

  Logger.log('🔍 Testando conexão Jira...');
  Logger.log(`   URL: ${baseUrl || '❌ não configurada'}`);
  Logger.log(`   Email: ${email || '❌ não configurado'}`);
  Logger.log(`   Token: ${token ? '✅ ' + token.length + ' caracteres' : '❌ não configurado'}`);

  if (!baseUrl || !email || !token) { Logger.log('Configure as Script Properties.'); return; }

  const sessionCookie = getJiraSession(baseUrl, email, token);
  const url  = `${baseUrl}/rest/api/2/search`;
  const body = JSON.stringify({ jql: JQL, maxResults: 1, fields: ['summary', FIELD_MUNICIPIO, FIELD_VERTICAL] });
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'Cookie': sessionCookie, 'Accept': 'application/json', 'Content-Type': 'application/json' },
    payload: body,
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  if (code === 200) {
    const data = JSON.parse(resp.getContentText());
    Logger.log(`✅ Conexão OK! Total de issues no filtro: ${data.total}`);
    if (data.issues.length > 0) {
      const i = data.issues[0];
      const v = i.fields[FIELD_VERTICAL];
      Logger.log(`   Exemplo — ${i.key}`);
      Logger.log(`   Município: ${i.fields[FIELD_MUNICIPIO] || '(vazio)'}`);
      Logger.log(`   Vertical:  ${v ? v.value : '(vazio)'}`);
    }
  } else {
    Logger.log(`❌ Erro HTTP ${code}: ${resp.getContentText().substring(0, 200)}`);
  }
}
