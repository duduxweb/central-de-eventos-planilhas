/** ===== AgendasTools.gs =====
 * Listagem/validação/preenchimento de agendas + mudar agenda + criar eventos pedindo agenda
 *
 * Requer:
 *  - resolverAgenda_(valorAgenda) em Código.gs (v2.1+) para resolver {cal,id,nome}
 *  - localizarEventoEmQualquerAgenda_(eventId) em Código.gs (para detectar agenda real pelo ID)
 */

/* =========================================
 * 0) Helpers genéricos (compartilhados)
 * ========================================= */

function encontrarColunaPorHeader_(sheet, nomesPossiveis) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const targets = (nomesPossiveis || []).map(s => String(s).toLowerCase());
  for (let c = 0; c < headers.length; c++) {
    const h = (headers[c] || '').toString().trim().toLowerCase();
    if (targets.includes(h)) return c + 1; // 1-based
  }
  return 0;
}

function encontrarColunaAgenda_(sheet) {
  return encontrarColunaPorHeader_(sheet, ['agenda']) || 0;
}

function carregarMapaAgendas_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Agendas');
  const map = {};
  if (!sh) return map;
  const last = sh.getLastRow();
  if (last <= 1) return map;
  const dados = sh.getRange(2, 1, last - 1, 2).getValues(); // A..B
  dados.forEach(([nome, id]) => {
    const n = (nome || '').toString().trim().toLowerCase();
    const i = (id || '').toString().trim();
    if (n && i) map[n] = i;
  });
  return map;
}

function checarAgendaEstrita_(val, mapAgendas) {
  try { if (CalendarApp.getCalendarById(val)) return true; } catch(_){}
  const maybeId = mapAgendas[(val || '').toString().trim().toLowerCase()];
  if (maybeId) { try { if (CalendarApp.getCalendarById(maybeId)) return true; } catch(_){ } }
  return false;
}

function columnToLetter_(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function escapeHtml_(s){
  return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}

function construirHtmlSelecaoAgendas_(callbackName) {
  const items = obterAgendasDaGuia_(); // [{nome, id}]
  const options = items.map(i => `<option value="${escapeHtml_(i.id)}">${escapeHtml_(i.nome)} — ${escapeHtml_(i.id)}</option>`).join('');
  return `
<!doctype html>
<html>
  <head><meta charset="utf-8"><style>
    body{font:14px Arial,sans-serif;padding:12px}
    label{display:block;margin-bottom:6px}
    select,button{width:100%;padding:8px;margin:6px 0}
    .muted{color:#666;font-size:12px}
  </style></head>
  <body>
    <h3>Selecionar agenda</h3>
    <label for="agenda">Escolha na lista (guia "Agendas"):</label>
    <select id="agenda">${options}</select>
    <button onclick="confirmar()">Confirmar</button>
    <p class="muted">A agenda será aplicada à operação em andamento.</p>
    <script>
      function confirmar(){
        var id = document.getElementById('agenda').value;
        google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).${callbackName}(id);
      }
    </script>
  </body>
</html>`;
}

function obterAgendasDaGuia_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Agendas');
  const out = [];
  if (!sh) return out;
  const last = sh.getLastRow();
  if (last <= 1) return out;
  const dados = sh.getRange(2, 1, last - 1, 2).getValues(); // A..B
  dados.forEach(([nome, id]) => {
    nome = (nome || '').toString().trim();
    id   = (id   || '').toString().trim();
    if (nome && id) out.push({nome, id});
  });
  return out;
}

/* =========================================
 * 1) Atualizar guia "Agendas" (A: Nome, B: ID)
 * ========================================= */
function listarAgendas() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Agendas');
  if (!sh) sh = ss.insertSheet('Agendas');
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['Nome', 'ID']]);
  const cals = CalendarApp.getAllCalendars();
  const linhas = cals.map(c => [c.getName(), c.getId()]);
  if (linhas.length) sh.getRange(2, 1, linhas.length, 2).setValues(linhas);
  SpreadsheetApp.getUi().alert(`Agendas atualizadas: ${linhas.length}`);
}

/* =========================================
 * 2) Validar coluna "Agenda" no intervalo selecionado
 * ========================================= */
function validarAgendasNoIntervalo() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getActiveSheet();
  const sel = ss.getActiveRange();
  if (!sel) return SpreadsheetApp.getUi().alert('Selecione um intervalo contendo a coluna "Agenda".');

  const colAgenda = encontrarColunaAgenda_(sh) || 8; // fallback H
  const c0 = sel.getColumn(), c1 = c0 + sel.getNumColumns() - 1;
  if (colAgenda < c0 || colAgenda > c1) {
    const colLetter = columnToLetter_(colAgenda);
    SpreadsheetApp.getUi().alert(`A coluna "Agenda" (${colLetter}) não está dentro da seleção atual.`);
    return;
  }

  const r0 = Math.max(sel.getRow(), 2), r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) return SpreadsheetApp.getUi().alert('Nada para validar.');

  const rngValores = sh.getRange(r0, colAgenda, r1 - r0 + 1, 1);
  const valores = rngValores.getValues();
  const mapAgendas = carregarMapaAgendas_();

  const bg = [], note = [];
  let invalidos = 0;
  for (let i = 0; i < valores.length; i++) {
    const val = (valores[i][0] || '').toString().trim();
    if (!val) { bg.push(['']); note.push(['']); continue; }
    const ok = checarAgendaEstrita_(val, mapAgendas);
    if (ok) { bg.push(['']); note.push(['']); }
    else { bg.push(['#F8D7DA']); note.push(['Agenda não encontrada na guia "Agendas" nem como ID válido.']); invalidos++; }
  }

  rngValores.setBackgrounds(bg);
  rngValores.setNotes(note);
  SpreadsheetApp.getUi().alert(`Validação concluída.\nLinhas inválidas: ${invalidos}`);
}

/* =========================================
 * 3) Preencher coluna "Agenda" nas linhas selecionadas
 *    — AGORA BUSCA A AGENDA REAL PELO ID (coluna F)
 * ========================================= */
function preencherAgendaNasLinhasSelecionadas() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getActiveSheet();
  const sel = ss.getActiveRange();
  const ui  = SpreadsheetApp.getUi();
  if (!sel) return ui.alert('Selecione as LINHAS onde deseja preencher a coluna "Agenda".');

  const colAgenda = encontrarColunaAgenda_(sh) || 8; // H
  const colId     = encontrarColunaPorHeader_(sh, ['id','id do evento']) || 6; // F

  const c0 = sel.getColumn(), c1 = c0 + sel.getNumColumns() - 1;
  if (colAgenda < c0 || colAgenda > c1) {
    const colLetter = columnToLetter_(colAgenda);
    return ui.alert(`A coluna "Agenda" (${colLetter}) não está dentro da seleção atual.`);
  }

  const r0 = Math.max(sel.getRow(), 2), r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) return ui.alert('Nada para preencher.');

  const resp = ui.prompt(
    'Preencher coluna "Agenda"',
    'Digite "TUDO" para sobrescrever também as células já preenchidas.\n' +
    'Ou deixe em branco para preencher apenas as células vazias (recomendado).',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const overwriteAll = String(resp.getResponseText() || '').trim().toUpperCase() === 'TUDO';

  const rngAgenda = sh.getRange(r0, colAgenda, r1 - r0 + 1, 1);
  const rngId     = sh.getRange(r0, colId,     r1 - r0 + 1, 1);
  const agendas   = rngAgenda.getValues(); // [[H], ...]
  const ids       = rngId.getValues();     // [[F], ...]

  const out = [];
  let atualizadas = 0;

  for (let i = 0; i < agendas.length; i++) {
    const atualH = (agendas[i][0] || '').toString().trim();
    const idRaw  = (ids[i][0]      || '').toString().trim();

    // Se já tem valor em H e NÃO é "TUDO", mantemos
    if (atualH && !overwriteAll) {
      out.push([atualH]);
      continue;
    }

    // Se tem ID, buscamos em qual agenda ele existe
    if (idRaw && typeof localizarEventoEmQualquerAgenda_ === 'function') {
      try {
        const origem = localizarEventoEmQualquerAgenda_(idRaw); // { cal, id, nome } | null
        if (origem && origem.cal) {
          const label = origem.nome || origem.id;
          out.push([label]);
          if (label !== atualH) atualizadas++;
          continue;
        }
      } catch(_) { /* segue */ }
    }

    // Sem ID (ou não achou): NÃO forçamos agenda padrão; mantemos o que estava
    out.push([atualH]);
  }

  rngAgenda.setValues(out);
  ui.alert(`Preenchimento concluído.\nCélulas atualizadas: ${atualizadas}`);
}

/* =========================================
 * 4) Alterar agenda da(s) linha(s) selecionada(s) — com seletor
 * ========================================= */
function alterarAgendaLinhasSelecionadas() {
  const html = construirHtmlSelecaoAgendas_('__moverEventosSelecionadosParaDestino');
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle('Alterar agenda'));
}

function __moverEventosSelecionadosParaDestino(destinoId) {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getActiveSheet();
  const sel = ss.getActiveRange();
  const ui  = SpreadsheetApp.getUi();

  if (!destinoId) { ui.alert('Nenhuma agenda selecionada.'); return; }

  const colAgenda = encontrarColunaAgenda_(sh) || 8;
  const colId     = encontrarColunaPorHeader_(sh, ['id','id do evento']) || 6;
  const colStatus = encontrarColunaPorHeader_(sh, ['status']) || 7;
  const colTitulo = encontrarColunaPorHeader_(sh, ['titulo','título','title']) || 3;
  const colData   = encontrarColunaPorHeader_(sh, ['data','data início','data inicio']) || 1;
  const colLocal  = encontrarColunaPorHeader_(sh, ['local','location']) || 4;
  const colDesc   = encontrarColunaPorHeader_(sh, ['descricao','descrição','description']) || 5;

  if (!sel) { ui.alert('Selecione as LINHAS com os eventos que deseja mover.'); return; }
  const c0 = sel.getColumn(), c1 = c0 + sel.getNumColumns() - 1;
  if (colId < c0 || colId > c1 || colAgenda < c0 || colAgenda > c1) {
    ui.alert('A seleção precisa conter, ao menos, as colunas "ID" e "Agenda".');
    return;
  }

  const r0 = Math.max(sel.getRow(), 2), r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) { ui.alert('Nada para processar.'); return; }

  const calDestino = CalendarApp.getCalendarById(destinoId);
  const nomeDestino = calDestino ? calDestino.getName() : destinoId;

  const apiAtiva = (function(){ try { return typeof Calendar !== 'undefined' && Calendar.Events && Calendar.Events.move; } catch(_){ return false; } })();
  let ok = 0, falhas = 0;

  for (let r = r0; r <= r1; r++) {
    try {
      const idEventoRaw = (sh.getRange(r, colId).getValue() || '').toString().trim();
      if (!idEventoRaw) { if (colStatus) sh.getRange(r, colStatus).setValue('Sem ID.'); continue; }

      const agendaAtualValor = (sh.getRange(r, colAgenda).getValue() || '').toString().trim();
      const origem = resolverAgenda_(agendaAtualValor);

      if (apiAtiva) {
        const eventIdForApi = idEventoRaw.endsWith('@google.com') ? idEventoRaw.slice(0, -12) : idEventoRaw;
        Calendar.Events.move(origem.id, eventIdForApi, destinoId, { sendUpdates:'all', supportsAttachments:true });
        sh.getRange(r, colAgenda).setValue(nomeDestino || destinoId);
        if (colStatus) sh.getRange(r, colStatus).setValue(`Movido (API) p/ ${nomeDestino} em ${formatarDataCompleta(new Date())}`);
      } else {
        const ev = origem.cal.getEventById(idEventoRaw);
        if (!ev) throw new Error('Evento não encontrado na origem.');
        const titulo = (sh.getRange(r, colTitulo).getValue() || ev.getTitle() || '').toString();
        const local  = (sh.getRange(r, colLocal).getValue()  || ev.getLocation() || '').toString();
        const desc   = (sh.getRange(r, colDesc).getValue()   || ev.getDescription() || '').toString();
        let novo;
        if (ev.isAllDayEvent()) {
          const dataPlan = sh.getRange(r, colData).getValue();
          const data = dataPlan instanceof Date ? dataPlan : ev.getAllDayStartDate();
          novo = calDestino.createAllDayEvent(titulo, data, { location: local, description: desc });
        } else {
          novo = calDestino.createEvent(titulo, ev.getStartTime(), ev.getEndTime(), { location: local, description: desc });
        }
        ev.deleteEvent();
        sh.getRange(r, colId).setValue(novo.getId());
        sh.getRange(r, colAgenda).setValue(nomeDestino || destinoId);
        if (colStatus) sh.getRange(r, colStatus).setValue(`Movido (fallback) p/ ${nomeDestino} em ${formatarDataCompleta(new Date())}`);
      }
      ok++; Utilities.sleep(80);
    } catch (e) {
      if (colStatus) sh.getRange(r, colStatus).setValue('Erro ao mover: ' + e.message);
      falhas++;
    }
  }

  ui.alert(`Concluído. Movidos: ${ok} | Falhas: ${falhas}`);
}

/* =========================================
 * 5) Criar evento(s) nas linhas selecionadas SEM ID pedindo agenda
 * ========================================= */
function criarEventosComSelecaoDeAgenda() {
  const html = construirHtmlSelecaoAgendas_('__criarEventosSelecionadosParaDestino');
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle('Criar eventos (sem ID)'));
}

function __criarEventosSelecionadosParaDestino(destinoId) {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getActiveSheet();
  const sel = ss.getActiveRange();
  const ui  = SpreadsheetApp.getUi();

  if (!destinoId) { ui.alert('Nenhuma agenda selecionada.'); return; }
  if (!sel) { ui.alert('Selecione as LINHAS com eventos a criar (sem ID na coluna F).'); return; }

  // Colunas padrão do layout "Eventos"
  const colData   = encontrarColunaPorHeader_(sh, ['data','data início','data inicio']) || 1; // A
  const colTitulo = encontrarColunaPorHeader_(sh, ['titulo','título','title']) || 3;          // C
  const colLocal  = encontrarColunaPorHeader_(sh, ['local','location']) || 4;                 // D
  const colDesc   = encontrarColunaPorHeader_(sh, ['descricao','descrição','description']) || 5; // E
  const colId     = encontrarColunaPorHeader_(sh, ['id','id do evento']) || 6;               // F
  const colStatus = encontrarColunaPorHeader_(sh, ['status']) || 7;                           // G
  const colAgenda = encontrarColunaAgenda_(sh) || 8;                                          // H

  const r0 = Math.max(sel.getRow(), 2), r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) { ui.alert('Nada para processar (provavelmente selecionou apenas o cabeçalho).'); return; }

  const calDestino = CalendarApp.getCalendarById(destinoId);
  const nomeDestino = calDestino ? calDestino.getName() : destinoId;

  let criados = 0, pulados = 0, erros = 0;

  for (let r = r0; r <= r1; r++) {
    try {
      const id = (sh.getRange(r, colId).getValue() || '').toString().trim();
      if (id) { pulados++; continue; } // só cria onde NÃO tem ID

      const dataVal = sh.getRange(r, colData).getValue();
      const titulo  = (sh.getRange(r, colTitulo).getValue() || '').toString().trim();
      const local   = (sh.getRange(r, colLocal).getValue()  || '').toString().trim();
      const desc    = (sh.getRange(r, colDesc).getValue()   || '').toString().trim();

      if (!titulo || !(dataVal instanceof Date)) {
        if (colStatus) sh.getRange(r, colStatus).setValue('Faltando título ou data.');
        continue;
      }

      // cria como evento de dia inteiro (mesma regra do seu v1)
      const novo = calDestino.createAllDayEvent(titulo, dataVal, { location: local, description: desc });

      // grava ID, H (agenda) e status
      sh.getRange(r, colId).setValue(novo.getId());
      sh.getRange(r, colAgenda).setValue(nomeDestino || destinoId);
      if (colStatus) sh.getRange(r, colStatus).setValue(`Criado em: ${formatarDataCompleta(new Date())}`);

      criados++;
      Utilities.sleep(60);
    } catch (e) {
      if (colStatus) sh.getRange(r, colStatus).setValue('Erro ao criar: ' + e.message);
      erros++;
    }
  }

  ui.alert(`Concluído. Criados: ${criados} | Pulados (já tinham ID): ${pulados} | Erros: ${erros}`);
}

/* =========================================
 * 6) APLICAR DROPDOWN NA COLUNA H (corrigido)
 * ========================================= */
function aplicarListaAgendasNaColunaH() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getActiveSheet();
  const sel = ss.getActiveRange();
  const ui  = SpreadsheetApp.getUi();

  if (!sel) {
    ui.alert('Selecione as LINHAS onde deseja aplicar o dropdown na coluna H.');
    return;
  }

  // Sempre aplicar na coluna H (8)
  const COL_H = 8;
  const c0 = sel.getColumn(), c1 = c0 + sel.getNumColumns() - 1;
  if (COL_H < c0 || COL_H > c1) {
    const colLetter = columnToLetter_(COL_H);
    ui.alert(`A coluna ${colLetter} (Agenda) não está dentro da seleção atual.\nSelecione colunas que incluam a coluna H.`);
    return;
  }

  // Linhas (ignora cabeçalho)
  const r0 = Math.max(sel.getRow(), 2);
  const r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) {
    ui.alert('Nada para aplicar (provavelmente selecionou apenas o cabeçalho).');
    return;
  }

  // Guia Agendas deve existir
  const shAg = ss.getSheetByName('Agendas');
  if (!shAg) {
    ui.alert('Guia "Agendas" não encontrada. Use "Atualizar lista de agendas" antes.');
    return;
  }

  const lastAg = shAg.getLastRow();
  if (lastAg <= 1) {
    ui.alert('A guia "Agendas" não tem itens (A2:A vazio). Atualize a lista de agendas.');
    return;
  }

  // Fonte da lista: Agendas!A2:A (Nomes)
  const rangeLista = shAg.getRange('A2:A');

  // Validação correta: requireValueInRange
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeLista, true) // true = mostra dropdown
    .setAllowInvalid(true)                 // mantenha true para não travar valores prévios fora da lista
    .build();

  // Aplica na coluna H das linhas selecionadas
  sh.getRange(r0, COL_H, r1 - r0 + 1, 1).setDataValidation(validation);

  ui.alert('Dropdown de agendas aplicado na coluna H (usando Agendas!A2:A).\nSe aparecer aviso, escolha um valor da lista e sincronize.');
}
