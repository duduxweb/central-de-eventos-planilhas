/** ===== Utils.gs =====
 * PONTO ÚNICO de criação de menus.
 * Mantém os menus da planilha SEMPRE neste arquivo.
 *
 * Observações:
 * - GARANTA que NENHUM outro arquivo contenha função onOpen().
 * - Este arquivo só cria menus e utilidades leves (ex.: limpar status).
 */

const CONFIGURACOES_ABA = 'Configurações';
const EVENTOS_SHEET     = 'Eventos';

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Descobre a "agenda atual" (se houver função para isso)
  const agendaAtualInfo = _safeGetAgendaAtualInfo_(); // {nome, id}
  const labelAgendaAtual = `• Agenda atual: ${agendaAtualInfo.nome}`;

  /* =========================
   * MENU 1 — 📅 Gerenciar Eventos
   * ========================= */
  ui.createMenu('📅 Gerenciar Eventos')
    .addItem(labelAgendaAtual, 'mostrarAgendaAtual')
    .addSubMenu(
      ui.createMenu('🗓️ Escolher agenda')
        .addItem('Usar agenda da linha selecionada em "Agendas"', 'selecionarAgendaPelaLinha')
        .addItem('Digitar ID de calendário…', 'selecionarAgendaPorPrompt')
        .addItem('Usar Configurações!B1', 'selecionarAgendaPorConfiguracoes')
        .addItem('Usar calendário padrão', 'selecionarAgendaPadrao')
        .addSeparator()
        .addItem('Ver agenda atual', 'mostrarAgendaAtual')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Guia atual')
        .addItem('Sincronizar TODOS (guia atual)', 'sincronizarTodosGuiaAtual')
        .addItem('Sincronizar SELECIONADOS (guia atual)', 'sincronizarSelecionadosGuiaAtual')
        .addItem('Limpar Status (guia atual)', 'limparStatusGuiaAtual')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Aba "Eventos" (legado)')
        .addItem('Sincronizar Todos', 'sincronizarTodosEventos')
        .addItem('Sincronizar Selecionados', 'sincronizarEventosSelecionados')
        .addItem('Apagar Eventos Selecionados', 'apagarEventosSelecionados')
        .addItem('Limpar Status (aba "Eventos")', 'limparStatus')
    )
    .addSeparator()
    .addItem('Inicializar Calendário (v1)', 'initCalendario')
    .addItem('Apagar Eventos Futuros (v1 / ID_CALENDARIO)', 'apagarTodosEventos')
    .addToUi();

  /* =========================
   * MENU 2 — 📎 Anexos & Agendas
   * ========================= */
  ui.createMenu('📎 Anexos & Agendas')
    .addItem('Listar todas as agendas', 'listarAgendas')
    .addSeparator()
    .addItem('Criar página "Anexos" (modelo)', 'criarAbaAnexosSeNaoExistir')
    .addItem('Sincronizar anexos (linhas selecionadas)', 'sincronizarAnexosSelecionados')
    .addItem('Mover eventos selecionados para outra agenda…', 'moverEventosSelecionadosParaOutraAgenda')
    .addSeparator()
    .addItem('Testar permissão por agenda…', 'testarPermissaoAgendaPrompt')
    .addItem('Testar permissões (guia Agendas)', 'testarPermissoesGuiaAgendas')
    .addToUi();
}

/** Limpa a coluna de Status da GUIA ATUAL (detecta coluna por cabeçalho; fallback G). */
function limparStatusGuiaAtual() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh) return;
  const last = sh.getLastRow();
  if (last <= 1) {
    SpreadsheetApp.getUi().alert('Nada para limpar nesta guia.');
    return;
  }
  const col = (typeof _colStatus_ === 'function' ? _colStatus_(sh) : 7) || 7; // fallback G
  sh.getRange(2, col, last - 1, 1).clearContent();
  SpreadsheetApp.getUi().alert('Coluna de status limpa (guia atual)!');
}

/** Limpa a coluna de Status (G) da aba "Eventos" (compat). */
function limparStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(EVENTOS_SHEET);
  if (!aba) {
    SpreadsheetApp.getUi().alert(`Aba "${EVENTOS_SHEET}" não encontrada.`);
    return;
  }
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha <= 1) {
    SpreadsheetApp.getUi().alert('Nada para limpar.');
    return;
  }
  aba.getRange(2, 7, ultimaLinha - 1, 1).clearContent();
  SpreadsheetApp.getUi().alert('Coluna de status limpa na aba "Eventos"!');
}

/** Formatação de data padrão do projeto (compartilhada com Código.gs). */
function formatarDataCompleta(data) {
  const ano     = data.getFullYear();
  const mes     = String(data.getMonth() + 1).padStart(2, '0');
  const dia     = String(data.getDate()).padStart(2, '0');
  const horas   = String(data.getHours()).padStart(2, '0');
  const minutos = String(data.getMinutes()).padStart(2, '0');
  const segundos= String(data.getSeconds()).padStart(2, '0');
  return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
}

/* =========================
 * Helpers internos locais
 * ========================= */
function mostrarAgendaAtual() {
  const ui = SpreadsheetApp.getUi();
  const info = _safeGetAgendaAtualInfo_();
  ui.alert(
    'Agenda atual',
    `Nome: ${info.nome}\nID: ${info.id || '(indefinido)'}\n\nDica: use "🗓️ Escolher agenda" para alterar.`,
    ui.ButtonSet.OK
  );
}

/** Tenta obter a agenda "atual" a partir da função global getAgendaOperacaoPadrao_().
 * Se não existir, tenta Configurações!B1; se falhar, usa o calendário padrão.
 */
function _safeGetAgendaAtualInfo_() {
  try {
    if (typeof getAgendaOperacaoPadrao_ === 'function') {
      const op = getAgendaOperacaoPadrao_(); // {cal, id, nome}
      if (op && (op.nome || op.id)) return { nome: op.nome || op.id, id: op.id || '' };
    }
  } catch (_) {}

  // fallback: Configurações!B1
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName(CONFIGURACOES_ABA);
    const id = cfg ? (cfg.getRange('B1').getValue() || '').toString().trim() : '';
    if (id) {
      try {
        const cal = CalendarApp.getCalendarById(id);
        if (cal) return { nome: cal.getName(), id: cal.getId() };
      } catch (_) {}
      return { nome: id, id: id };
    }
  } catch (_) {}

  // último fallback: calendário padrão
  try {
    const cal = CalendarApp.getDefaultCalendar();
    return { nome: cal.getName(), id: cal.getId() };
  } catch (_) {
    return { nome: '(desconhecida)', id: '' };
  }
}
