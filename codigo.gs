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

// JQL para implantações pendentes e em andamento
const JQL_IMPLANTACOES = `(labels not in (implantaçãoRecusada) OR labels is EMPTY) AND issuetype = Implantação AND "Equipe responsável" not in (Revenda, Parceiros, Produto, "Produto extensões", Tribunais) AND resolution = Unresolved AND status not in ("Produto contratado", Reprovada) AND (Município in ("Abdon Batista", Agrolândia, "Anita Garibaldi", Angelina, Anchieta, "Balneário Arroio do Silva", "Balneário Barra do Sul", "Balneário Camboriú", "Balneário Piçarras", Bandeirante, "Barra Bonita", "Barra Velha", "Bela Vista do Toldo", Belmonte, "Benedito Novo", Brunópolis, Caçador, Calmon, "Campo Alegre", "Capão Alto", Chapecó, Concórdia, "Dona Emma", "Erval Velho", Ermo, "Frei Rogério", Iraceminha, Imbuia, Ipira, Ipuaçu, Itá, Itajaí, Jupiá, Lacerdópolis, "Lajeado Grande", "Leoberto Leal", "Lindóia do Sul", "Luiz Alves", Luzerna, Mafra, Massaranduba, Meleiro, Modelo, "Morro da Fumaça", "Morro Grande", Penha, Peritiba, "Pescaria Brava", Pomerode, "Praia Grande", "Rio do Sul", "Rio Fortuna", "Rio Rufino", Saltinho, "Santa Terezinha", "São Bernardino", "São Bonifácio", "São Cristovão do Sul", "São João do Oeste", "São José do Cedro", "São Martinho", "São Miguel da Boa Vista", "São Pedro de Alcântara", Tangará, "Treze de Maio", Tigrinhos, Timbó, Treviso, Videira) OR Município in ("Campos Novos") AND Entidade = "CIMPLASC - CONSORCIO INTERMUNICIPAL DE SANEAMENTO BASICO MEIO AMBIENTE ATENCAO A SANIDADE DOS PRODUTOS DE ORIGEM AGROPECUARIA SEGURANCA ALIMENTAR - Campos Novos/SC") ORDER BY status DESC, cf[21500] DESC, issuetype ASC, Município ASC, cf[10300] ASC, cf[22902] ASC, assignee DESC`;

const FIELD_MUNICIPIO = 'customfield_10331'; // Município (string)
const FIELD_VERTICAL  = 'customfield_10300'; // Vertical  ({ value: "Saúde" })
const SHEET_TAB_NAME  = 'Jira_Chamados';
const IMPL_TAB_NAME   = 'Jira_Implantacoes';
const CND_TAB_NAME    = 'CND_Municipios';
const CND_SHEET_ID             = '16axvbTygJCmXY2zT2FL3a5BYDNrUz-tIwTTrifkwwcQ';
const NPS_TAB_NAME             = 'NPS_Calculado';
const COLABORADORES_SHEET_ID   = '1ksgbwdf5dgsoI9XUiEobFKzsytA_XaOFSNUDlOX0Apk';
const COLABORADORES_GID        = 1645653528;
const COLABORADORES_TAB_NAME   = 'Colaboradores';
const PAGE_SIZE       = 100;
const SLEEP_MS        = 200;

// ────────────────────────────────────────────────────────────
// ENTRY POINT — disparado automaticamente pelo trigger (6h)
// ────────────────────────────────────────────────────────────
function onTimeTrigger() {
  const inicio = new Date();
  Logger.log(`▶ Iniciando coleta: ${inicio.toLocaleString('pt-BR')}`);
  try {
    const issues = fetchJiraIssues();
    Logger.log(`  Issues coletadas: ${issues.length}`);
    const rows = aggregateByMunicipioVertical(issues);
    writeJiraChamados(rows);
    Logger.log(`  Jira: ${rows.length} linhas gravadas`);
    fetchAndStoreCND();
    fetchAndStoreNPS();
    fetchAndStoreColaboradores();
    fetchAndStoreImplantacoes();
    Logger.log(`✅ Concluído em ${Math.round((new Date()-inicio)/1000)}s`);
  } catch (e) {
    Logger.log(`❌ ERRO: ${e.message}`);
    throw e;
  }
}

// ────────────────────────────────────────────────────────────
// CND — busca dados do endpoint CND e salva na planilha
// Roda server-side com auth do proprietário do script
// ────────────────────────────────────────────────────────────
function fetchAndStoreCND() {
  Logger.log('  Buscando dados CND da planilha...');
  const ssCND  = SpreadsheetApp.openById(CND_SHEET_ID);
  const tabCND = ssCND.getSheets()[0]; // primeira aba
  const values = tabCND.getDataRange().getValues();

  if (values.length < 2) { Logger.log('  CND: planilha vazia'); return; }

  const headers = values[0];

  // Detectar colunas por nome (busca parcial, case-insensitive)
  const col = (termo) => headers.findIndex(h =>
    h.toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').includes(termo));

  const iMun  = col('munic');
  const iPort = col('portf');
  const iP1   = headers.findIndex(h => /per[^\d]*1/i.test(h));
  const iP2   = headers.findIndex(h => /per[^\d]*2/i.test(h));
  const iP3   = headers.findIndex(h => /per[^\d]*3/i.test(h));

  Logger.log(`  Colunas detectadas — Município:${iMun} Portfólio:${iPort} P1:${iP1} P2:${iP2} P3:${iP3}`);

  const filtrados = values.slice(1).filter(row => {
    const port = iPort >= 0 ? (row[iPort] || '').toString().toLowerCase() : '';
    return port.includes('pequenas') && row[iMun];
  });
  Logger.log(`  CND: ${filtrados.length} registros de Pequenas Contas`);

  const ts = new Date().toISOString();
  const writeRows = filtrados.map(row => [
    row[iMun]  || '',
    iPort >= 0 ? (row[iPort] || '') : '',
    iP1   >= 0 ? (row[iP1]  || '') : '',
    iP2   >= 0 ? (row[iP2]  || '') : '',
    iP3   >= 0 ? (row[iP3]  || '') : '',
    ts,
  ]);

  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  let tab       = ss.getSheetByName(CND_TAB_NAME);
  if (!tab) { tab = ss.insertSheet(CND_TAB_NAME); }

  tab.getRange(1, 1, 1, 6).setValues([['municipio','portfolio','periodo1','periodo2','periodo3','atualizado_em']]);
  if (tab.getLastRow() > 1) tab.getRange(2, 1, tab.getLastRow() - 1, 6).clearContent();
  if (writeRows.length > 0) tab.getRange(2, 1, writeRows.length, 6).setValues(writeRows);

  const h = tab.getRange(1, 1, 1, 6);
  h.setBackground('#1E3A5F'); h.setFontColor('#FFFFFF'); h.setFontWeight('bold');
  tab.setFrozenRows(1);
  Logger.log(`  CND_Municipios: ${writeRows.length} linhas gravadas`);
}

// Diagnóstico: testa leitura CND sem gravar
function testarCND() {
  Logger.log('🔍 Testando leitura CND...');
  const ssCND  = SpreadsheetApp.openById(CND_SHEET_ID);
  const tab    = ssCND.getSheets()[0];
  const values = tab.getDataRange().getValues();
  Logger.log(`  Total de linhas (com header): ${values.length}`);
  if (values.length > 0) Logger.log(`  Headers: ${values[0].join(' | ')}`);
  const pequenas = values.slice(1).filter(r => (r[1] || '').toString().toLowerCase().includes('pequenas'));
  Logger.log(`  Linhas de Pequenas Contas: ${pequenas.length}`);
  if (pequenas.length > 0) Logger.log(`  Exemplo: ${JSON.stringify(pequenas[0])}`);
}

// ────────────────────────────────────────────────────────────
// NPS — agrega comentarios_NPS por município e salva NPS_Calculado
// Promotores 9-10, Neutros 7-8, Detratores 0-6
// NPS = % Promotores − % Detratores (resultado de -100 a 100)
// ────────────────────────────────────────────────────────────
function fetchAndStoreNPS() {
  Logger.log('  Calculando NPS de comentarios_NPS...');
  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);

  const tab = ss.getSheetByName('comentarios_NPS');
  if (!tab) { Logger.log('  NPS: aba comentarios_NPS não encontrada'); return; }

  const values = tab.getDataRange().getValues();
  if (values.length < 2) { Logger.log('  NPS: planilha vazia'); return; }

  const headers = values[0];

  // Detectar colunas por nome
  const iMun   = headers.findIndex(h => h.toString().toLowerCase().includes('municipio'));
  const iScore = headers.findIndex(h => /npsgeralemail|npsgeral/i.test(h));

  Logger.log(`  Colunas NPS — Município:${iMun} Score:${iScore}`);
  if (iMun < 0 || iScore < 0) { Logger.log('  NPS: colunas não encontradas'); return; }

  // Agregar por município
  const map = {};
  values.slice(1).forEach(function(row) {
    const mun   = (row[iMun] || '').toString().trim();
    const score = parseFloat(row[iScore]);
    if (!mun || isNaN(score)) return;
    if (!map[mun]) map[mun] = { total: 0, promotores: 0, neutros: 0, detratores: 0 };
    map[mun].total++;
    if      (score >= 9) map[mun].promotores++;
    else if (score >= 7) map[mun].neutros++;
    else                 map[mun].detratores++;
  });

  const ts   = new Date().toISOString();
  const rows = Object.keys(map).sort().map(function(mun) {
    const d   = map[mun];
    const nps = d.total > 0 ? Math.round((d.promotores / d.total) * 100 - (d.detratores / d.total) * 100) : 0;
    return [mun, d.total, d.promotores, d.neutros, d.detratores, nps, ts];
  });

  let tabOut = ss.getSheetByName(NPS_TAB_NAME);
  if (!tabOut) tabOut = ss.insertSheet(NPS_TAB_NAME);

  tabOut.getRange(1, 1, 1, 7).setValues([['municipio','total','promotores','neutros','detratores','nps_score','atualizado_em']]);
  if (tabOut.getLastRow() > 1) tabOut.getRange(2, 1, tabOut.getLastRow() - 1, 7).clearContent();
  if (rows.length > 0) tabOut.getRange(2, 1, rows.length, 7).setValues(rows);

  const h = tabOut.getRange(1, 1, 1, 7);
  h.setBackground('#1E3A5F'); h.setFontColor('#FFFFFF'); h.setFontWeight('bold');
  tabOut.setFrozenRows(1);
  Logger.log(`  NPS_Calculado: ${rows.length} municípios gravados`);
}

// ────────────────────────────────────────────────────────────
// COLABORADORES — lê planilha de colaboradores e salva na aba Colaboradores
// ────────────────────────────────────────────────────────────
function fetchAndStoreColaboradores() {
  Logger.log('  Buscando dados de Colaboradores...');
  const ssColabs = SpreadsheetApp.openById(COLABORADORES_SHEET_ID);
  const tab = ssColabs.getSheets().find(s => s.getSheetId() === COLABORADORES_GID);
  if (!tab) {
    Logger.log('  Colaboradores: aba gid=' + COLABORADORES_GID + ' não encontrada');
    return;
  }

  const values = tab.getDataRange().getValues();
  if (values.length < 2) { Logger.log('  Colaboradores: sem dados'); return; }

  const headers = values[0];
  var iAnalista = -1, iVaga = -1, iArea = -1, iRegiao = -1, iAtend = -1;
  headers.forEach(function(h, i) {
    var s = h.toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (s.includes('analista'))           iAnalista = i;
    if (s === 'vaga')                     iVaga = i;
    if (s.includes('area de atuacao'))    iArea = i;
    if (s.includes('regiao'))             iRegiao = i;
    if (s.includes('area de atendimento')) iAtend = i;
  });

  var rows = values.slice(1)
    .filter(function(row) { return row.some(function(c) { return c !== ''; }); })
    .map(function(row) {
      return [
        iAnalista >= 0 ? String(row[iAnalista] || '') : '',
        iVaga     >= 0 ? String(row[iVaga]     || '') : '',
        iArea     >= 0 ? String(row[iArea]     || '') : '',
        iRegiao   >= 0 ? String(row[iRegiao]   || '') : '',
        iAtend    >= 0 ? String(row[iAtend]    || '') : '',
      ];
    });

  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  var tabOut = ss.getSheetByName(COLABORADORES_TAB_NAME);
  if (!tabOut) tabOut = ss.insertSheet(COLABORADORES_TAB_NAME);

  tabOut.getRange(1, 1, 1, 5).setValues([['analista', 'vaga', 'area_de_atuacao', 'regiao', 'area_de_atendimento']]);
  if (tabOut.getLastRow() > 1) tabOut.getRange(2, 1, tabOut.getLastRow() - 1, 5).clearContent();
  if (rows.length > 0) tabOut.getRange(2, 1, rows.length, 5).setValues(rows);

  const hdr = tabOut.getRange(1, 1, 1, 5);
  hdr.setBackground('#1E3A5F'); hdr.setFontColor('#FFFFFF'); hdr.setFontWeight('bold');
  tabOut.setFrozenRows(1);
  Logger.log('  Colaboradores: ' + rows.length + ' registros gravados');
}

// ────────────────────────────────────────────────────────────
// IMPLANTAÇÕES — coleta issues de implantação do Jira e salva pivot
// Agrupa por Município × Vertical → aba Jira_Implantacoes
// ────────────────────────────────────────────────────────────
function fetchAndStoreImplantacoes() {
  Logger.log('  Buscando implantações do Jira...');
  const props    = PropertiesService.getScriptProperties();
  const baseUrl  = props.getProperty('JIRA_BASE_URL') || '';
  const email    = props.getProperty('JIRA_EMAIL')    || '';
  const password = props.getProperty('JIRA_API_TOKEN') || '';

  if (!baseUrl || !email || !password) {
    Logger.log('  Implantações: Script Properties não configuradas, pulando.');
    return;
  }

  const sessionCookie = getJiraSession(baseUrl, email, password);
  const headers = { 'Cookie': sessionCookie, 'Accept': 'application/json', 'Content-Type': 'application/json' };
  const fields  = ['summary', FIELD_MUNICIPIO, FIELD_VERTICAL];

  const allIssues = [];
  let startAt = 0;

  while (true) {
    const url  = `${baseUrl}/rest/api/2/search`;
    const body = JSON.stringify({ jql: JQL_IMPLANTACOES, fields: fields, maxResults: PAGE_SIZE, startAt: startAt });

    const resp = UrlFetchApp.fetch(url, { method: 'post', headers, payload: body, muteHttpExceptions: true });
    const code = resp.getResponseCode();

    if (code !== 200) {
      Logger.log(`  Implantações: Jira HTTP ${code}: ${resp.getContentText().substring(0, 200)}`);
      return;
    }

    const data = JSON.parse(resp.getContentText());
    allIssues.push(...data.issues);
    Logger.log(`  Impl pág ${Math.floor(startAt/PAGE_SIZE)+1}: ${data.issues.length} issues (${allIssues.length}/${data.total})`);

    if (allIssues.length >= data.total || data.issues.length === 0) break;
    startAt += PAGE_SIZE;
    Utilities.sleep(SLEEP_MS);
  }

  Logger.log(`  Implantações coletadas: ${allIssues.length}`);

  // Agregar por Município × Vertical
  const map = {};
  const ts  = new Date().toISOString();
  allIssues.forEach(issue => {
    const municipio = (issue.fields[FIELD_MUNICIPIO] || 'Não informado').toString().trim();
    const vObj      = issue.fields[FIELD_VERTICAL];
    const vertical  = (vObj && vObj.value) ? vObj.value.trim() : 'Não informado';
    const key       = `${municipio}||${vertical}`;
    map[key]        = (map[key] || 0) + 1;
  });

  const rows = Object.entries(map).map(([key, count]) => {
    const [municipio, vertical] = key.split('||');
    return [municipio, vertical, count, ts];
  }).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));

  // Gravar na aba Jira_Implantacoes
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  let tab       = ss.getSheetByName(IMPL_TAB_NAME);
  if (!tab) { tab = ss.insertSheet(IMPL_TAB_NAME); Logger.log(`  Aba "${IMPL_TAB_NAME}" criada.`); }

  tab.getRange(1, 1, 1, 4).setValues([['municipio','vertical','total','atualizado_em']]);
  const last = tab.getLastRow();
  if (last > 1) tab.getRange(2, 1, last - 1, 4).clearContent();
  if (rows.length > 0) tab.getRange(2, 1, rows.length, 4).setValues(rows);

  const h = tab.getRange(1, 1, 1, 4);
  h.setBackground('#1E3A5F'); h.setFontColor('#FFFFFF'); h.setFontWeight('bold');
  tab.setFrozenRows(1);
  Logger.log(`  "${IMPL_TAB_NAME}": ${rows.length} linhas gravadas (${allIssues.length} issues, ${Object.keys(map).length} combinações mun×vertical).`);
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
