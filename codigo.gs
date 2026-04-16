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
// Abas que possuem versão histórica (_Hist)
const HIST_ENABLED = ['Jira_Chamados', 'Jira_Implantacoes', 'NPS_Calculado', 'CND_Municipios', 'CND_Federal', 'CND_Estadual', 'Risco de Exclusão', 'Colaboradores'];

function doGet(e) {
  const callback  = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  const sheetName = (e && e.parameter && e.parameter.sheet)    ? e.parameter.sheet    : 'Jira_Chamados';
  const dateParam = (e && e.parameter && e.parameter.date)     ? e.parameter.date     : null;

  let payload;
  try {
    const today      = new Date().toISOString().slice(0, 10);
    const useHist    = dateParam && dateParam !== today && HIST_ENABLED.includes(sheetName);
    const targetTab  = useHist ? sheetName + '_Hist' : sheetName;
    const filterDate = useHist ? dateParam : null;

    const data = readSheetData(targetTab, filterDate);
    payload = JSON.stringify({ status: 'ok', sheet: sheetName, date: dateParam, data: data });
  } catch (err) {
    payload = JSON.stringify({ status: 'error', message: err.message });
  }

  const output = callback
    ? ContentService.createTextOutput(`${callback}(${payload})`).setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);

  return output;
}

// Lê uma aba e retorna array de objetos com os cabeçalhos como chaves.
// filterDate (opcional): filtra apenas linhas cujo campo atualizado_em começa com essa data (YYYY-MM-DD).
function readSheetData(sheetName, filterDate) {
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

  const iDate = filterDate ? headers.indexOf('atualizado_em') : -1;

  return values.slice(1)
    .filter(row => {
      if (!row.some(cell => cell !== '' && cell !== null)) return false;
      if (filterDate && iDate >= 0) {
        const d = row[iDate] ? row[iDate].toString().slice(0, 10) : '';
        return d === filterDate;
      }
      return true;
    })
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
const FIELD_PRAZO     = 'customfield_25801'; // Prazo contratual da implantação (date "YYYY-MM-DD")
const SHEET_TAB_NAME  = 'Jira_Chamados';
const IMPL_TAB_NAME         = 'Jira_Implantacoes';
const IMPL_ISSUES_TAB_NAME  = 'Jira_Implantacoes_Issues';
const CND_TAB_NAME    = 'CND_Municipios';
const CND_SHEET_ID             = '16axvbTygJCmXY2zT2FL3a5BYDNrUz-tIwTTrifkwwcQ';
// CND Federal (SICONFI) e Estadual (e-Sfinge SC) — planilhas mantidas pelo governo
const SICONFI_SHEET_ID  = '1vrRNrQoKhFllqH8OIxQveeUFt-LUbdw1hsBwrh5OAI4';
const SICONFI_GID       = 1585345281;
const ESFINGE_SHEET_ID  = '1hRXUjAvwJKhTecn0SYDT2n5_EomjNexStlmNfKOcVvs';
const ESFINGE_GID       = 1585345281;
const CND_FEDERAL_TAB   = 'CND_Federal';
const CND_ESTADUAL_TAB  = 'CND_Estadual';
const NPS_TAB_NAME             = 'NPS_Calculado';
const COLABORADORES_SHEET_ID   = '1ksgbwdf5dgsoI9XUiEobFKzsytA_XaOFSNUDlOX0Apk';
const COLABORADORES_GID        = 1645653528;
const COLABORADORES_TAB_NAME   = 'Colaboradores';
const PAGE_SIZE            = 100;
const SLEEP_MS             = 200;
const HIST_RETENTION_DAYS  = 90;  // dias de retenção nas abas _Hist

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
    fetchAndStoreCNDFederal();
    fetchAndStoreCNDEstadual();
    fetchAndStoreNPS();
    fetchAndStoreColaboradores();
    snapshotRiscoExclusaoHistory();
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

  // Histórico diário
  const dateStrCND = new Date().toISOString().slice(0, 10);
  appendToHistory(ss, CND_TAB_NAME + '_Hist',
    ['municipio', 'portfolio', 'periodo1', 'periodo2', 'periodo3', 'atualizado_em'],
    writeRows, 5, dateStrCND);
}

// ────────────────────────────────────────────────────────────
// Helpers compartilhados por CND Federal/Estadual
// ────────────────────────────────────────────────────────────
function _normStr(s) {
  return (s === null || s === undefined ? '' : s.toString()).trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}
function _readSheetByGid(sheetId, gid) {
  const ss  = SpreadsheetApp.openById(sheetId);
  const tab = ss.getSheets().find(s => s.getSheetId() === gid);
  if (!tab) throw new Error('Aba gid=' + gid + ' não encontrada em ' + sheetId);
  return tab.getDataRange().getValues();
}
// Filtro compartilhado: apenas Betha Sistemas + Canal Pequenas e Médias Contas
function _isPequenasContasBetha(row, iPrest, iCanal) {
  return _normStr(row[iPrest]) === 'betha sistemas'
      && _normStr(row[iCanal]).indexOf('pequenas') >= 0
      && _normStr(row[iCanal]).indexOf('medias') >= 0;
}

// Whitelist de municípios do portfólio de Pequenas Contas
// (normalizados: sem acentos, minúsculos, sem espaços duplos)
const PORTFOLIO_MUNICIPIOS = [
  'abdon batista','agrolandia','anchieta','angelina','anita garibaldi',
  'balneario arroio do silva','balneario barra do sul','balneario camboriu','balneario picarras',
  'bandeirante','barra bonita','barra velha','bela vista do toldo','belmonte','benedito novo',
  'brunopolis','cacador','calmon','campo alegre','capao alto','chapeco','concordia',
  'dona emma','ermo','erval velho','frei rogerio','imbuia','ipira','ipuacu','iraceminha','ita',
  'itajai','jupia','lacerdopolis','lajeado grande','leoberto leal','lindoia do sul','luiz alves',
  'luzerna','mafra','massaranduba','meleiro','modelo','morro da fumaca','morro grande','penha',
  'peritiba','pescaria brava','pomerode','praia grande','rio do sul','rio fortuna','rio rufino',
  'saltinho','santa terezinha','sao bernardino','sao bonifacio','sao cristovao do sul',
  'sao joao do oeste','sao jose do cedro','sao martinho','sao miguel da boa vista',
  'sao pedro de alcantara','tangara','tigrinhos','timbo','treviso','treze de maio','videira'
];
const PORTFOLIO_MUNICIPIOS_SET = (function() {
  const s = {};
  PORTFOLIO_MUNICIPIOS.forEach(m => { s[m] = true; });
  return s;
})();
function _isMunicipioPortfolio(nome) {
  return !!PORTFOLIO_MUNICIPIOS_SET[_normStr(nome)];
}

// Formata nome de município em Title Case com preposições em minúsculo
// Ex.: "SÃO JOÃO DO OESTE" → "São João do Oeste"
function _titleCaseMun(s) {
  if (s === null || s === undefined) return '';
  const lower = { 'de':1,'da':1,'do':1,'dos':1,'das':1,'e':1,'a':1,'o':1,'à':1 };
  return s.toString().toLowerCase().split(/\s+/).filter(Boolean).map((w, i) => {
    if (i > 0 && lower[w]) return w;
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(' ');
}

// Extrai data de atualização da fonte da célula A1 das planilhas CND (SICONFI/e-Sfinge)
// Formata como dd/MM/yyyy HH:mm:ss. A1 pode ser um Date (Google Sheets) ou string.
function _extractStatusFonte(values) {
  if (!values || !values.length || !values[0] || !values[0].length) return '';
  const a1 = values[0][0];
  const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';

  // 1) Se for objeto Date → formata direto (só data)
  if (a1 instanceof Date && !isNaN(a1.getTime())) {
    return Utilities.formatDate(a1, tz, 'dd/MM/yyyy');
  }

  // 2) Se for string → extrai DD/MM/YYYY
  const s = (a1 === null || a1 === undefined ? '' : a1.toString());
  const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    const dd = m[1].padStart(2,'0');
    const mm = m[2].padStart(2,'0');
    const yy = m[3].length === 2 ? '20' + m[3] : m[3];
    return dd + '/' + mm + '/' + yy;
  }
  return '';
}

// ────────────────────────────────────────────────────────────
// CND FEDERAL (SICONFI) — status mensal por município (12 meses do ano)
// Cabeçalhos estão na linha 3 da planilha de origem.
// Grava CND_Federal: municipio, tipo, jan..dez, atualizado_em
// ────────────────────────────────────────────────────────────
function fetchAndStoreCNDFederal() {
  Logger.log('  Buscando CND Federal (SICONFI)...');
  const values = _readSheetByGid(SICONFI_SHEET_ID, SICONFI_GID);
  if (values.length < 4) { Logger.log('  SICONFI: sem dados'); return; }

  const headerRow = values[2]; // linha 3 = cabeçalhos
  const idx = (name) => headerRow.findIndex(h => _normStr(h) === _normStr(name));
  const iPrest = idx('Prestador');
  const iCanal = idx('Canal');
  const iMun   = idx('Municipio');
  const iTipo  = idx('Tipo');
  const meses  = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const iMes   = meses.map(m => idx(m));

  if (iPrest < 0 || iCanal < 0 || iMun < 0) {
    Logger.log('  SICONFI: cabeçalhos Prestador/Canal/Municipio não encontrados');
    return;
  }

  // Filtra por Prestador = 'Betha Sistemas' + Canal = 'Pequenas e Médias Contas' + município do portfólio
  const filtrados = values.slice(3).filter(r =>
    _isPequenasContasBetha(r, iPrest, iCanal) && r[iMun] && _isMunicipioPortfolio(r[iMun]));
  Logger.log('  SICONFI: ' + filtrados.length + ' linhas após filtro (Betha + Pequenas e Médias + portfólio)');

  const nowIso  = new Date().toISOString();
  const dateStr = nowIso.slice(0, 10);
  const statusFonte = _extractStatusFonte(values);
  Logger.log('  SICONFI: status fonte (A1) = ' + statusFonte);

  const writeRows = filtrados.map(r => {
    const out = [
      _titleCaseMun(r[iMun]),
      iTipo >= 0 ? String(r[iTipo] || '').trim() : '',
      statusFonte,
    ];
    for (let i = 0; i < 12; i++) out.push(iMes[i] >= 0 ? String(r[iMes[i]] || '').trim() : '');
    out.push(nowIso);
    return out;
  });

  const headers = ['municipio','tipo','status_fonte','jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez','atualizado_em'];
  const W = headers.length;

  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  let tab = ss.getSheetByName(CND_FEDERAL_TAB);
  if (!tab) tab = ss.insertSheet(CND_FEDERAL_TAB);

  tab.getRange(1, 1, 1, W).setValues([headers]);
  if (tab.getLastRow() > 1) tab.getRange(2, 1, tab.getLastRow() - 1, W).clearContent();
  if (writeRows.length > 0) tab.getRange(2, 1, writeRows.length, W).setValues(writeRows);

  const h = tab.getRange(1, 1, 1, W);
  h.setBackground('#1E3A5F'); h.setFontColor('#FFFFFF'); h.setFontWeight('bold');
  tab.setFrozenRows(1);
  Logger.log('  CND_Federal: ' + writeRows.length + ' linhas gravadas');

  appendToHistory(ss, CND_FEDERAL_TAB + '_Hist', headers, writeRows, W - 1, dateStr);
}

// ────────────────────────────────────────────────────────────
// CND ESTADUAL (e-Sfinge SC) — 3 certidões quadrimestrais (Sim/Não)
// Cabeçalhos estão na linha 3. Colunas "Certidão ..." usadas como P1/P2/P3.
// Grava CND_Estadual: municipio, entidade, meses_atraso, tipo_atraso,
//   periodo1, periodo2, periodo3, p1_label, p2_label, p3_label, atualizado_em
// ────────────────────────────────────────────────────────────
function fetchAndStoreCNDEstadual() {
  Logger.log('  Buscando CND Estadual (e-Sfinge SC)...');
  const values = _readSheetByGid(ESFINGE_SHEET_ID, ESFINGE_GID);
  if (values.length < 4) { Logger.log('  e-Sfinge: sem dados'); return; }

  const headerRow = values[2];
  const idx = (name) => headerRow.findIndex(h => _normStr(h) === _normStr(name));
  const iPrest = idx('Prestador');
  const iCanal = idx('Canal');
  const iMun   = idx('Municipio');
  const iEnt   = idx('Entidade');
  const iMesesAtraso = idx('Meses em atraso');
  const iTipoAtraso  = idx('Tipo atraso');

  // Detecta as 3 colunas "Certidão ..." dinamicamente, preservando o texto do rótulo
  const certIdx = [];
  const certLabels = [];
  headerRow.forEach((h, i) => {
    const s = h ? h.toString().trim() : '';
    if (/^Certid[ãa]o\s/i.test(s)) {
      certIdx.push(i);
      certLabels.push(s.replace(/^Certid[ãa]o\s+/i, '').trim());
    }
  });
  Logger.log('  e-Sfinge: certidões detectadas = ' + certLabels.join(' | '));

  if (iPrest < 0 || iCanal < 0 || iMun < 0 || certIdx.length < 3) {
    Logger.log('  e-Sfinge: cabeçalhos obrigatórios ausentes (certIdx=' + certIdx.length + ')');
    return;
  }

  // Filtra por Prestador = 'Betha Sistemas' + Canal = 'Pequenas e Médias Contas' + município do portfólio
  const filtrados = values.slice(3).filter(r =>
    _isPequenasContasBetha(r, iPrest, iCanal) && r[iMun] && _isMunicipioPortfolio(r[iMun]));
  Logger.log('  e-Sfinge: ' + filtrados.length + ' linhas após filtro (Betha + Pequenas e Médias + portfólio)');

  const normCert = (v) => {
    const s = _normStr(v);
    if (s === 'sim') return 'Com CND';
    if (s === 'nao') return 'Sem CND';
    return '';
  };

  const nowIso  = new Date().toISOString();
  const dateStr = nowIso.slice(0, 10);
  const statusFonte = _extractStatusFonte(values);
  Logger.log('  e-Sfinge: status fonte (A1) = ' + statusFonte);

  const writeRows = filtrados.map(r => [
    _titleCaseMun(r[iMun]),
    iEnt >= 0 ? _titleCaseMun(r[iEnt]) : '',
    iMesesAtraso >= 0 ? String(r[iMesesAtraso] || '').trim() : '',
    iTipoAtraso  >= 0 ? String(r[iTipoAtraso]  || '').trim() : '',
    statusFonte,
    normCert(r[certIdx[0]]),
    normCert(r[certIdx[1]]),
    normCert(r[certIdx[2]]),
    certLabels[0] || '',
    certLabels[1] || '',
    certLabels[2] || '',
    nowIso,
  ]);

  const headers = ['municipio','entidade','meses_atraso','tipo_atraso','status_fonte','periodo1','periodo2','periodo3','p1_label','p2_label','p3_label','atualizado_em'];
  const W = headers.length;

  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  let tab = ss.getSheetByName(CND_ESTADUAL_TAB);
  if (!tab) tab = ss.insertSheet(CND_ESTADUAL_TAB);

  tab.getRange(1, 1, 1, W).setValues([headers]);
  if (tab.getLastRow() > 1) tab.getRange(2, 1, tab.getLastRow() - 1, W).clearContent();
  if (writeRows.length > 0) tab.getRange(2, 1, writeRows.length, W).setValues(writeRows);

  const h = tab.getRange(1, 1, 1, W);
  h.setBackground('#1E3A5F'); h.setFontColor('#FFFFFF'); h.setFontWeight('bold');
  tab.setFrozenRows(1);
  Logger.log('  CND_Estadual: ' + writeRows.length + ' linhas gravadas');

  appendToHistory(ss, CND_ESTADUAL_TAB + '_Hist', headers, writeRows, W - 1, dateStr);
}

// Diagnóstico: lê SICONFI direto da planilha e loga um município específico
// Uso: alterar o parâmetro e executar no editor Apps Script → ver Logs
function diagnosticarSICONFI(municipioAlvo) {
  municipioAlvo = municipioAlvo || 'Ipira';
  Logger.log('🔍 Diagnóstico SICONFI — alvo: ' + municipioAlvo);
  const values = _readSheetByGid(SICONFI_SHEET_ID, SICONFI_GID);
  Logger.log('  Total linhas: ' + values.length);
  Logger.log('  Cabeçalho (linha 3): ' + values[2].join(' | '));
  const headerRow = values[2];
  const idx = (name) => headerRow.findIndex(h => _normStr(h) === _normStr(name));
  const iPrest = idx('Prestador'), iCanal = idx('Canal'), iMun = idx('Municipio'), iTipo = idx('Tipo');
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const iMes = meses.map(m => idx(m));
  Logger.log('  Índices: Prestador=' + iPrest + ' Canal=' + iCanal + ' Mun=' + iMun + ' Tipo=' + iTipo);
  Logger.log('  Índices meses: ' + iMes.join(','));
  const matches = values.slice(3).filter(r =>
    r[iMun] && _normStr(r[iMun]) === _normStr(municipioAlvo));
  Logger.log('  Linhas encontradas para ' + municipioAlvo + ': ' + matches.length);
  matches.forEach((r, i) => {
    Logger.log('  Linha ' + (i+1) + ' — Prestador="' + r[iPrest] + '" Canal="' + r[iCanal] + '" Tipo="' + r[iTipo] + '"');
    meses.forEach((m, j) => {
      Logger.log('    ' + m + ' (col ' + iMes[j] + ') = "' + r[iMes[j]] + '"');
    });
  });
}

// Diagnóstico e-Sfinge: mostra todas as linhas de um município na planilha fonte
function diagnosticarEsfinge(municipioAlvo) {
  municipioAlvo = municipioAlvo || 'Agrolandia';
  Logger.log('🔍 Diagnóstico e-Sfinge — alvo: ' + municipioAlvo);
  const values = _readSheetByGid(ESFINGE_SHEET_ID, ESFINGE_GID);
  Logger.log('  Total linhas: ' + values.length);
  const headerRow = values[2];
  Logger.log('  Cabeçalho (linha 3): ' + headerRow.join(' | '));
  const idx = (name) => headerRow.findIndex(h => _normStr(h) === _normStr(name));
  const iPrest = idx('Prestador'), iCanal = idx('Canal'), iMun = idx('Municipio'), iEnt = idx('Entidade');
  Logger.log('  Índices: Prestador=' + iPrest + ' Canal=' + iCanal + ' Mun=' + iMun + ' Entidade=' + iEnt);
  const matches = values.slice(3).filter(r =>
    r[iMun] && _normStr(r[iMun]).indexOf(_normStr(municipioAlvo)) >= 0);
  Logger.log('  Linhas encontradas para "' + municipioAlvo + '": ' + matches.length);
  matches.forEach((r, i) => {
    const certs = headerRow.map((h, j) => /^Certid/i.test(h ? h.toString() : '') ? h + '=' + r[j] : null).filter(Boolean);
    Logger.log('  Linha ' + (i+1) + ': Prestador="' + r[iPrest] + '" | Canal="' + r[iCanal] + '" | Entidade="' + r[iEnt] + '" | ' + certs.join(' | '));
    Logger.log('    → passa filtro Betha+Pequenas: ' + _isPequenasContasBetha(r, iPrest, iCanal));
    Logger.log('    → no portfólio: ' + _isMunicipioPortfolio(r[iMun]));
  });
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

  // Histórico diário
  const dateStrNPS = new Date().toISOString().slice(0, 10);
  appendToHistory(ss, NPS_TAB_NAME + '_Hist',
    ['municipio', 'total', 'promotores', 'neutros', 'detratores', 'nps_score', 'atualizado_em'],
    rows, 6, dateStrNPS);
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

  const nowIso  = new Date().toISOString();
  const dateStr = nowIso.slice(0, 10);

  var rows = values.slice(1)
    .filter(function(row) { return row.some(function(c) { return c !== ''; }); })
    .map(function(row) {
      return [
        iAnalista >= 0 ? String(row[iAnalista] || '') : '',
        iVaga     >= 0 ? String(row[iVaga]     || '') : '',
        iArea     >= 0 ? String(row[iArea]     || '') : '',
        iRegiao   >= 0 ? String(row[iRegiao]   || '') : '',
        iAtend    >= 0 ? String(row[iAtend]    || '') : '',
        nowIso,
      ];
    });

  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  var tabOut = ss.getSheetByName(COLABORADORES_TAB_NAME);
  if (!tabOut) tabOut = ss.insertSheet(COLABORADORES_TAB_NAME);

  tabOut.getRange(1, 1, 1, 6).setValues([['analista', 'vaga', 'area_de_atuacao', 'regiao', 'area_de_atendimento', 'atualizado_em']]);
  if (tabOut.getLastRow() > 1) tabOut.getRange(2, 1, tabOut.getLastRow() - 1, 6).clearContent();
  if (rows.length > 0) tabOut.getRange(2, 1, rows.length, 6).setValues(rows);

  const hdr = tabOut.getRange(1, 1, 1, 6);
  hdr.setBackground('#1E3A5F'); hdr.setFontColor('#FFFFFF'); hdr.setFontWeight('bold');
  tabOut.setFrozenRows(1);
  Logger.log('  Colaboradores: ' + rows.length + ' registros gravados');

  appendToHistory(ss, COLABORADORES_TAB_NAME + '_Hist',
    ['analista', 'vaga', 'area_de_atuacao', 'regiao', 'area_de_atendimento', 'atualizado_em'],
    rows, 5, dateStr);
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
  const fields  = ['summary', FIELD_MUNICIPIO, FIELD_VERTICAL, FIELD_PRAZO, 'status'];

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

  // Agregar por Município × Vertical — inclui status de prazo
  const map  = {};
  const ts   = new Date().toISOString();
  const hoje = new Date().toISOString().slice(0, 10); // "YYYY-MM-DD"

  allIssues.forEach(issue => {
    const municipio = (issue.fields[FIELD_MUNICIPIO] || 'Não informado').toString().trim();
    const vObj      = issue.fields[FIELD_VERTICAL];
    const vertical  = (vObj && vObj.value) ? vObj.value.trim() : 'Não informado';
    const prazo     = issue.fields[FIELD_PRAZO] || null; // "YYYY-MM-DD" ou null
    const key       = `${municipio}||${vertical}`;

    if (!map[key]) map[key] = { total: 0, atrasados: 0, no_prazo: 0, sem_prazo: 0, prazo_minimo: null };
    map[key].total++;

    if (!prazo) {
      map[key].sem_prazo++;
    } else if (prazo < hoje) {
      map[key].atrasados++;
      // Guarda o prazo vencido mais antigo (mais crítico)
      if (!map[key].prazo_minimo || prazo < map[key].prazo_minimo) {
        map[key].prazo_minimo = prazo;
      }
    } else {
      map[key].no_prazo++;
    }
  });

  const rows = Object.entries(map).map(([key, d]) => {
    const [municipio, vertical] = key.split('||');
    return [municipio, vertical, d.total, d.atrasados, d.no_prazo, d.sem_prazo, d.prazo_minimo || '', ts];
  }).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));

  // Gravar na aba Jira_Implantacoes
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  let tab       = ss.getSheetByName(IMPL_TAB_NAME);
  if (!tab) { tab = ss.insertSheet(IMPL_TAB_NAME); Logger.log(`  Aba "${IMPL_TAB_NAME}" criada.`); }

  const header = [['municipio','vertical','total','atrasados','no_prazo','sem_prazo','prazo_minimo','atualizado_em']];
  tab.getRange(1, 1, 1, 8).setValues(header);
  const last = tab.getLastRow();
  if (last > 1) tab.getRange(2, 1, last - 1, 8).clearContent();
  if (rows.length > 0) tab.getRange(2, 1, rows.length, 8).setValues(rows);

  const h = tab.getRange(1, 1, 1, 8);
  h.setBackground('#1E3A5F'); h.setFontColor('#FFFFFF'); h.setFontWeight('bold');
  tab.setFrozenRows(1);

  const totalAtrasados = Object.values(map).reduce((s, d) => s + d.atrasados, 0);
  const totalNoPrazo   = Object.values(map).reduce((s, d) => s + d.no_prazo, 0);
  const totalSemPrazo  = Object.values(map).reduce((s, d) => s + d.sem_prazo, 0);
  Logger.log(`  "${IMPL_TAB_NAME}": ${rows.length} linhas (${allIssues.length} issues — ${totalAtrasados} atrasadas, ${totalNoPrazo} no prazo, ${totalSemPrazo} sem prazo).`);

  // Gravar issues individuais na aba Jira_Implantacoes_Issues
  let issuesTab = ss.getSheetByName(IMPL_ISSUES_TAB_NAME);
  if (!issuesTab) { issuesTab = ss.insertSheet(IMPL_ISSUES_TAB_NAME); Logger.log(`  Aba "${IMPL_ISSUES_TAB_NAME}" criada.`); }

  const issuesHeader = [['key','url','summary','status','municipio','vertical','prazo','atualizado_em']];
  issuesTab.getRange(1, 1, 1, 8).setValues(issuesHeader);
  const lastIssue = issuesTab.getLastRow();
  if (lastIssue > 1) issuesTab.getRange(2, 1, lastIssue - 1, 8).clearContent();

  const issueRows = allIssues.map(issue => {
    const municipio = (issue.fields[FIELD_MUNICIPIO] || 'Não informado').toString().trim();
    const vObj      = issue.fields[FIELD_VERTICAL];
    const vertical  = (vObj && vObj.value) ? vObj.value.trim() : 'Não informado';
    const prazo     = issue.fields[FIELD_PRAZO] || '';
    const status    = (issue.fields.status && issue.fields.status.name) ? issue.fields.status.name : '';
    return [issue.key, `${baseUrl}/browse/${issue.key}`, issue.fields.summary || '', status, municipio, vertical, prazo, ts];
  }).sort((a, b) => a[4].localeCompare(b[4], 'pt-BR') || a[5].localeCompare(b[5], 'pt-BR'));

  if (issueRows.length > 0) issuesTab.getRange(2, 1, issueRows.length, 8).setValues(issueRows);
  const ih = issuesTab.getRange(1, 1, 1, 8);
  ih.setBackground('#1E3A5F'); ih.setFontColor('#FFFFFF'); ih.setFontWeight('bold');
  issuesTab.setFrozenRows(1);
  Logger.log(`  "${IMPL_ISSUES_TAB_NAME}": ${issueRows.length} issues individuais gravados.`);

  // Histórico diário (pivot agregada apenas — issues individuais não vão para histórico)
  const dateStrImpl = new Date().toISOString().slice(0, 10);
  appendToHistory(ss, IMPL_TAB_NAME + '_Hist',
    ['municipio', 'vertical', 'total', 'atrasados', 'no_prazo', 'sem_prazo', 'prazo_minimo', 'atualizado_em'],
    rows, 7, dateStrImpl);
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

  // Histórico diário
  const dateStr = new Date().toISOString().slice(0, 10);
  appendToHistory(ss, SHEET_TAB_NAME + '_Hist',
    ['municipio', 'vertical', 'total_chamados', 'atualizado_em'],
    rows, 3, dateStr);
}

// ────────────────────────────────────────────────────────────
// HISTÓRICO — appenda snapshot diário e mantém 90 dias
// Parâmetros:
//   ss           → Spreadsheet da planilha principal
//   histTabName  → ex: 'Jira_Chamados_Hist'
//   headerRow    → array com nomes das colunas (ex: ['municipio','vertical','total_chamados','atualizado_em'])
//   newRows      → array de arrays com os dados novos
//   dateColIdx   → índice (0-based) da coluna de timestamp (atualizado_em)
//   dateStr      → 'YYYY-MM-DD' do dia atual
// ────────────────────────────────────────────────────────────
function appendToHistory(ss, histTabName, headerRow, newRows, dateColIdx, dateStr) {
  let histTab = ss.getSheetByName(histTabName);
  if (!histTab) {
    histTab = ss.insertSheet(histTabName);
    histTab.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    const hdr = histTab.getRange(1, 1, 1, headerRow.length);
    hdr.setBackground('#1E3A5F'); hdr.setFontColor('#FFFFFF'); hdr.setFontWeight('bold');
    histTab.setFrozenRows(1);
  }

  // Data de corte para retenção (90 dias atrás)
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - HIST_RETENTION_DAYS);
  const cutoffStr = cutoff.toISOString().slice(0, 10);

  // Ler linhas existentes, descartar as de hoje (serão substituídas) e as muito antigas
  let keepRows = [];
  const lastRow = histTab.getLastRow();
  if (lastRow > 1) {
    const existing = histTab.getRange(2, 1, lastRow - 1, headerRow.length).getValues();
    keepRows = existing.filter(row => {
      const d = row[dateColIdx] ? row[dateColIdx].toString().slice(0, 10) : '';
      return d && d !== dateStr && d >= cutoffStr;
    });
  }

  // Reescrever: linhas preservadas + novas
  const allRows = keepRows.concat(newRows);
  if (lastRow > 1) histTab.getRange(2, 1, lastRow - 1, headerRow.length).clearContent();
  if (allRows.length > 0) histTab.getRange(2, 1, allRows.length, headerRow.length).setValues(allRows);

  Logger.log(`  "${histTabName}": ${newRows.length} novas linhas para ${dateStr} (total histórico: ${allRows.length}).`);
}

// ────────────────────────────────────────────────────────────
// RISCO DE EXCLUSÃO — snapshot diário (aba populada manualmente)
// Lê o estado atual da aba "Risco de Exclusão" e grava em "Risco de Exclusão_Hist"
// ────────────────────────────────────────────────────────────
function snapshotRiscoExclusaoHistory() {
  Logger.log('  Snapshot Risco de Exclusão...');
  const props   = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID') || SHEET_ID_DEFAULT;
  const ss      = SpreadsheetApp.openById(sheetId);
  const tab     = ss.getSheetByName('Risco de Exclusão');
  if (!tab) { Logger.log('  Risco: aba não encontrada'); return; }

  const values = tab.getDataRange().getValues();
  if (values.length < 2) { Logger.log('  Risco: sem dados'); return; }

  const headers = values[0].map(function(h) {
    return h.toString().trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '_');
  });
  const iMun = headers.indexOf('municipio');
  const iRis = headers.indexOf('risco');
  const iMot = headers.indexOf('motivo');
  const iCon = headers.indexOf('consultor');

  const nowIso  = new Date().toISOString();
  const dateStr = nowIso.slice(0, 10);

  const rows = values.slice(1)
    .filter(function(r) { return r.some(function(c) { return c !== '' && c !== null; }); })
    .map(function(r) {
      return [
        iMun >= 0 ? String(r[iMun] || '') : '',
        iRis >= 0 ? String(r[iRis] || '') : '',
        iMot >= 0 ? String(r[iMot] || '') : '',
        iCon >= 0 ? String(r[iCon] || '') : '',
        nowIso,
      ];
    });

  appendToHistory(ss, 'Risco de Exclusão_Hist',
    ['municipio', 'risco', 'motivo', 'consultor', 'atualizado_em'],
    rows, 4, dateStr);
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
