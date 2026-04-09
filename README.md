# Portfólio Pequenas Contas — Dashboard Executivo

Dashboard executivo para monitoramento do portfólio de pequenas contas municipais de Santa Catarina. Exibe 6 visões analíticas com dados atualizados automaticamente.

## Visões disponíveis

| # | Visão | Fonte |
|---|-------|-------|
| 1 | Chamados em aberto por cliente | Jira |
| 2 | Chamados por cliente e vertical | Jira |
| 3 | Clientes com e sem CND | Sistema CND |
| 4 | Clientes com risco de exclusão | Google Sheets |
| 5 | NPS do portfólio | Google Sheets |
| 6 | NPS por cliente | Google Sheets |

---

## Setup inicial (executar uma única vez)

### Passo 1 — Publicar a planilha do Google Sheets

A planilha precisa ser pública para o dashboard ler os dados.

1. Abra a planilha: https://docs.google.com/spreadsheets/d/1MKsApbL7IPf5jAsAO9N03AxAcC3ptzYbrgNQrOr_R4s
2. Menu **Arquivo → Compartilhar → Publicar na web**
3. Selecione **Planilha inteira** → Formato **CSV** → clique **Publicar**
4. Repita para cada aba (NPS, Risco de Exclusão, Jira_Chamados)

> A aba **Jira_Chamados** será criada automaticamente pelo Apps Script no Passo 3.

---

### Passo 2 — Gerar token da API Jira

1. Acesse: https://id.atlassian.com/manage/api-tokens
2. Clique em **Criar token de API**
3. Dê o nome `portfolio-dashboard` e copie o token gerado

---

### Passo 3 — Configurar e instalar o Apps Script

1. Abra a planilha do Google Sheets
2. Menu **Extensões → Apps Script**
3. Apague todo o conteúdo do editor e cole o conteúdo do arquivo `codigo.gs`
4. Salve (Ctrl+S)

#### Configurar as propriedades do script

5. Menu lateral → **Configurações do projeto** (ícone de engrenagem)
6. Role até **Propriedades do script** → clique **Adicionar propriedade** e adicione:

| Propriedade | Valor |
|-------------|-------|
| `JIRA_BASE_URL` | `https://atendimento.betha.com.br` |
| `JIRA_EMAIL` | seu-email@betha.com.br |
| `JIRA_API_TOKEN` | token copiado no Passo 2 |
| `SHEET_ID` | `1MKsApbL7IPf5jAsAO9N03AxAcC3ptzYbrgNQrOr_R4s` |

#### Testar a conexão

7. No editor do Apps Script, selecione a função `testarConexaoJira` no dropdown
8. Clique **Executar**
9. Verifique os logs (View → Logs): deve aparecer `✅ Conexão OK! Total de issues no filtro: NNN`

#### Instalar o agendamento automático

10. Selecione a função `setupTrigger` no dropdown
11. Clique **Executar** (uma única vez)
12. O script passará a executar automaticamente a cada 6 horas

#### Executar manualmente a primeira vez

13. Selecione `onTimeTrigger` e clique **Executar**
14. Verifique na planilha: a aba `Jira_Chamados` deve ter sido criada com os dados

---

### Passo 4 — Ativar GitHub Pages

1. Faça push dos arquivos `index.html`, `codigo.gs` e `README.md` para o repositório
2. No GitHub: **Settings → Pages → Source: Deploy from a branch → main / (root)**
3. Aguarde ~60 segundos
4. Acesse: https://arimanoelgomes-ctrl.github.io/portfolio_pequenas_contas/

---

## Estrutura dos arquivos

```
portfolio_pequenas_contas/
├── index.html    ← Dashboard (HTML + CSS + JS em um único arquivo)
├── codigo.gs     ← Apps Script que coleta dados do Jira → Google Sheets
└── README.md     ← Este arquivo
```

---

## Estrutura das abas do Google Sheets

### `Jira_Chamados` — criada e mantida pelo Apps Script

| Coluna | Header | Descrição |
|--------|--------|-----------|
| A | `municipio` | Nome do município |
| B | `vertical` | Área (Saúde, Educação, etc.) |
| C | `total_chamados` | Número de chamados abertos |
| D | `atualizado_em` | Timestamp da última execução (ISO 8601) |

### `Risco de Exclusão` — mantida manualmente

| Coluna | Header | Valores aceitos |
|--------|--------|-----------------|
| A | `municipio` | Nome do município |
| B | `risco` | `Alto`, `Médio`, `Baixo`, `Sem Risco` |
| C | `motivo` | Texto livre (exibido no dashboard) |

### `NPS` — mantida manualmente

| Coluna | Header | Valores aceitos |
|--------|--------|-----------------|
| A | `municipio` | Nome do município |
| B | `nota` | Inteiro de 0 a 10 |
| C | `periodo` | Ex: `T1 2026` (texto livre) |

> **Importante:** Os cabeçalhos das colunas devem ser exatamente como mostrado acima (minúsculas, sem acentos onde indicado), pois o dashboard os usa para encontrar os dados.

---

## Revisões mensais (rotina recomendada)

Antes de cada apresentação ao gestor:

- [ ] Atualizar a aba **NPS** com as notas do período atual
- [ ] Revisar a aba **Risco de Exclusão** e atualizar status dos municípios
- [ ] Abrir o dashboard e clicar **⟳ Atualizar** para forçar recarregamento
- [ ] Verificar o timestamp "Jira:" no cabeçalho (deve ser recente)
- [ ] Navegar pelas 6 abas e validar os dados antes da reunião

---

## Diagnóstico de problemas

| Sintoma | Causa provável | Solução |
|---------|----------------|---------|
| Gráficos não carregam | Planilha não publicada | Repita o Passo 1 |
| "Dados indisponíveis" no Jira | Aba Jira_Chamados vazia | Execute `onTimeTrigger` manualmente |
| Erro 401 no Apps Script | Token Jira inválido | Gere novo token e atualize Script Properties |
| CND não carrega | Timeout ou erro de rede | Clique "Atualizar" e aguarde |
| Trigger parou de funcionar | Expirou autorização OAuth | Abra o Apps Script e re-execute `setupTrigger` |

---

## Informações técnicas

- **Frontend:** HTML + CSS + JavaScript puro, sem framework, sem build step
- **Gráficos:** [Chart.js 4.4](https://www.chartjs.org/) via CDN
- **Dados Jira:** Jira REST API v2 via Google Apps Script (autenticação Basic)
- **Dados CND:** JSONP via Apps Script publicado como web app
- **Dados Sheets:** Google Sheets publicado como CSV (`gviz/tq?tqx=out:csv`)
- **Campos Jira descobertos:** Município = `customfield_10331`, Vertical = `customfield_10300`
- **Atualização Jira:** Automática a cada 6 horas via trigger do Apps Script
