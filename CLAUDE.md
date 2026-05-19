# Portfólio Pequenas Contas — Premissas do Projeto

## ⚠️ PREMISSA CRÍTICA: Nunca quebrar leitura de dados históricos

Toda alteração de schema (adição, remoção ou reordenação de colunas) nas abas da
planilha Google Sheets **deve ser retrocompatível** com os dados históricos já
gravados nas abas `_Hist`.

### Abas com versão histórica (`HIST_ENABLED`)
```
Jira_Chamados | Jira_Chamados_Suporte | Jira_Implantacoes
NPS_Calculado | CND_Municipios | CND_Federal | CND_Estadual
Risco de Exclusão | Colaboradores
```

### Regras obrigatórias ao alterar o schema

1. **Novas colunas**: sempre adicionar ao **final** da linha de cabeçalho,
   nunca no meio — os dados históricos não terão a coluna e o `readSheetData`
   devolverá `undefined` para ela (aceitável; nunca quebrará a leitura).

2. **Remoção de colunas**: proibido nas abas `HIST_ENABLED`. Em abas de issues
   (`*_Issues`) que não têm `_Hist`, é permitido com cuidado.

3. **Reordenação de colunas**: proibido. O `readSheetData` lê por índice
   de cabeçalho, mas o `appendToHistory` e a dashboard dependem da ordem
   estável de algumas colunas-chave (`atualizado_em`, `municipio`, `vertical`).

4. **`appendToHistory`**: já possui detecção dinâmica do índice de
   `atualizado_em` para suportar schemas antigos. Ao adicionar colunas,
   confirmar que essa lógica continua funcional.

5. **Dashboard (`index.html`)**: ao consumir um campo novo, usar sempre
   acesso defensivo (`r.campo || ''` / `r.campo || '—'`) para que dados
   históricos sem o campo não causem erros.

6. **Antes de qualquer deploy**: verificar se a mudança afeta alguma aba
   `HIST_ENABLED` e, se sim, testar a leitura de uma data antiga no seletor
   de data da dashboard.

### Abas de issues (sem versão histórica — schema livre)
```
Jira_Chamados_Issues | Jira_Chamados_Suporte_Issues | Jira_Implantacoes_Issues
```
Essas abas são reescritas integralmente a cada coleta (`onTimeTrigger`),
portanto não têm dados históricos a preservar. Alterações de schema são
seguras desde que o Apps Script e o `index.html` sejam atualizados juntos.

---

## Stack do projeto

- **Backend**: Google Apps Script (`codigo.gs`) — Web App JSONP + coleta Jira
- **Frontend**: HTML/JS single-file (`index.html`) hospedado no GitHub Pages
  ou servido via Apps Script
- **Dados**: Google Sheets (`SHEET_ID = 1MKsApbL7IPf5jAsAO9N03AxAcC3ptzYbrgNQrOr_R4s`)
- **Fonte de issues**: Jira REST API v2, autenticado por session cookie
- **SLO field**: `customfield_24813`

## Fluxo de deploy

1. Editar `codigo.gs` → colar no Apps Script → executar `setupTrigger()` se
   alterou triggers, ou `onTimeTrigger()` para forçar coleta imediata.
2. Editar `index.html` → `git add && git commit && git push origin main`.
3. A dashboard lê os dados via JSONP da URL do Apps Script publicado.
