/** ===== Código.gs =====
 * Sincronismo de eventos com suporte à coluna "Agenda" (H),
 * detecção por cabeçalho em QUALQUER guia (para uso via menus),
 * fallback quando faltar Calendar.Events.move,
 * e mensagens de erro claras quando faltar PERMISSÃO (Action not allowed / Forbidden).
 *
 * >>> Importante: NÃO há ação automática ao editar a coluna H.
 * Use os MENUS para sincronizar (guia atual / seleção / aba "Eventos").
 *
 * Cabeçalho esperado (linha 1, nomes aproximados aceitos):
 * - Data (ou Data Início)
 * - (opcional) Data Fim  [não usada neste modelo]
 * - Título
 * - Local
 * - Descrição
 * - ID (ou ID do Evento)
 * - Status
 * - Agenda
 */

let ID_CALENDARIO = ''; // compatibilidade com v1 (usado em apagarTodosEventos)

/** Inicializa a variável ID_CALENDARIO a partir de Configurações!B1 (compatibilidade v1). */
function initCalendario() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaConfig = planilha.getSheetByName('Configurações');
  const id = abaConfig ? abaConfig.getRange('B1').getValue() : '';
  ID_CALENDARIO = id || SpreadsheetApp.getActive().getOwner().getEmail();
  Logger.log(`ID_CALENDARIO definido como: ${ID_CALENDARIO}`);
}

/** Resolve e retorna { cal, id, nome } a partir do valor da coluna Agenda. */
function resolverAgenda_(valorAgenda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const val = (valorAgenda || '').toString().trim();

  // 1) Tenta ID direto
  if (val) {
    try {
      const calDireto = CalendarApp.getCalendarById(val);
      if (calDireto) return { cal: calDireto, id: calDireto.getId(), nome: calDireto.getName() };
    } catch (_) {}
  }

  // 2) Tenta por NOME na aba "Agendas" (A: Nome, B: ID)
  if (val) {
    const sh = ss.getSheetByName('Agendas');
    if (sh) {
      const last = sh.getLastRow();
      if (last >= 2) {
        const lista = sh.getRange(2, 1, last - 1, 2).getValues(); // A..B
        for (let i = 0; i < lista.length; i++) {
          const nome = (lista[i][0] || '').toString().trim();
          const id   = (lista[i][1] || '').toString().trim();
          if (nome && id && nome.toLowerCase() === val.toLowerCase()) {
            const cal = CalendarApp.getCalendarById(id);
            if (cal) return { cal, id: cal.getId(), nome: cal.getName() };
          }
        }
      }
    }
  }

  // 3) Agenda de destino da operação (se oferecida por outro arquivo)
  if (typeof getAgendaOperacaoPadrao_ === 'function') {
    const op = getAgendaOperacaoPadrao_(); // { cal, id, nome } ou null
    if (op && op.cal) return op;
  }

  // 4) Fallback final: calendário padrão do usuário
  const calPadrao = CalendarApp.getDefaultCalendar();
  return { cal: calPadrao, id: calPadrao.getId(), nome: calPadrao.getName() };
}

/* =========================
 * Helpers de cabeçalho/colunas
 * ========================= */
function _headers_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || '').toString().trim());
}
function _findColByNames_(sheet, nomesPossiveis) {
  const headers = _headers_(sheet);
  const lowers = headers.map(h => h.toLowerCase());
  const targets = (nomesPossiveis || []).map(s => String(s).toLowerCase());
  const idx = lowers.findIndex(h => targets.includes(h));
  return idx >= 0 ? idx + 1 : 0; // 1-based
}
function _colAgenda_(sheet) { return _findColByNames_(sheet, ['agenda']); }
function _colId_(sheet) { return _findColByNames_(sheet, ['id', 'id do evento']); }
function _colDataInicio_(sheet) { return _findColByNames_(sheet, ['data', 'data início', 'data inicio']); }
function _colTitulo_(sheet) { return _findColByNames_(sheet, ['título', 'titulo', 'title']); }
function _colLocal_(sheet) { return _findColByNames_(sheet, ['local', 'location']); }
function _colDesc_(sheet) { return _findColByNames_(sheet, ['descrição', 'descricao', 'description']); }
function _colStatus_(sheet) { return _findColByNames_(sheet, ['status']); }

/* =========================
 * onEdit DESATIVADO (sem auto-sync)
 * ========================= */
function onEdit(e) {
  try {
    Logger.log('onEdit chamado, auto-sync desativado pelo usuário.');
  } catch (err) {
    Logger.log('onEdit erro: ' + err.message);
  }
}

/* =========================
 * Helpers de permissão / mensagens
 * ========================= */
function _isPermissionError_(err) {
  const m = (err && err.message ? String(err.message) : '').toLowerCase();
  return m.includes('action not allowed') ||
         m.includes('forbidden') ||
         m.includes('insufficient permission') ||
         m.includes('required permissions') ||
         m.includes('sem permissão');
}
function _msgSemPermissao_(acao, nomeOuId) {
  return `Sem permissão para ${acao} na agenda: ${nomeOuId}.
Verifique se sua conta tem "Fazer alterações nos eventos" nessa agenda.`;
}

/** Teste "silencioso" de escrita numa agenda (cria e apaga um evento de teste).
 * Retorna { ok:boolean, msg:string }.
 */
function testarPermissaoEscritaCalendario_(cal) {
  try {
    const hoje = new Date();
    const test = cal.createAllDayEvent('[_TESTE_PERMISSAO_]', hoje);
    try { test.deleteEvent(); } catch (_) {}
    return { ok: true, msg: 'OK: escrita permitida' };
  } catch (e) {
    if (_isPermissionError_(e)) {
      return { ok: false, msg: 'Sem permissão de escrita' };
    }
    return { ok: false, msg: 'Erro: ' + e.message };
  }
}

/* =========================
 * Sincronismo por GUIA (uso via menus)
 * ========================= */

/** Sincroniza UMA linha em QUALQUER sheet (detectando colunas por cabeçalho). */
function _sincronizarLinhaEmSheet_(sh, linha) {
  try {
    const cData   = _colDataInicio_(sh);
    const cTitulo = _colTitulo_(sh);
    const cLocal  = _colLocal_(sh);
    const cDesc   = _colDesc_(sh);
    const cId     = _colId_(sh);
    const cStatus = _colStatus_(sh);
    const cAgenda = _colAgenda_(sh);

    if (!cData || !cTitulo || !cId || !cAgenda) {
      if (cStatus) sh.getRange(linha, cStatus).setValue('Configuração de colunas incompleta (Data/Título/ID/Agenda).');
      return;
    }

    const dataInicioOriginal = sh.getRange(linha, cData).getValue();
    const titulo             = (sh.getRange(linha, cTitulo).getValue() || '').toString().trim();
    const local              = (sh.getRange(linha, cLocal).getValue()  || '').toString().trim();
    const descricao          = (sh.getRange(linha, cDesc).getValue()   || '').toString().trim();
    const idEvento           = (sh.getRange(linha, cId).getValue()     || '').toString().trim();
    const valorAgenda        = (sh.getRange(linha, cAgenda).getValue() || '').toString().trim();

    if (!titulo || !(dataInicioOriginal instanceof Date)) {
      if (cStatus) sh.getRange(linha, cStatus).setValue('Erro: faltando título ou data.');
      return;
    }

    const destino = resolverAgenda_(valorAgenda); // { cal, id, nome }
    const dataInicio = new Date(
      dataInicioOriginal.getFullYear(),
      dataInicioOriginal.getMonth(),
      dataInicioOriginal.getDate()
    );
    const agora = formatarDataCompleta(new Date());
    const nomeAgendaDestino = destino.nome || destino.id;

    if (!idEvento) {
      // ===== CRIAR =====
      try {
        const evento = destino.cal.createAllDayEvent(titulo, dataInicio, { location: local, description: descricao });
        sh.getRange(linha, cId).setValue(evento.getId());
        if (cStatus) sh.getRange(linha, cStatus).setValue(`Criado em: ${agora}`);
        sh.getRange(linha, cAgenda).setValue(nomeAgendaDestino);
      } catch (e) {
        if (cStatus) {
          const msg = _isPermissionError_(e) ? _msgSemPermissao_('criar', nomeAgendaDestino) : ('Erro ao criar: ' + e.message);
          sh.getRange(linha, cStatus).setValue(msg);
        }
      }
      return;
    }

    // ===== ATUALIZAR / MOVER =====

    // 1) Tenta atualizar direto na agenda de destino (H)
    try {
      let eventoDestino = destino.cal.getEventById(idEvento);
      if (eventoDestino) {
        try {
          eventoDestino.setTitle(titulo);
          eventoDestino.setLocation(local);
          eventoDestino.setDescription(descricao);
          eventoDestino.setAllDayDate(dataInicio);
          if (cStatus) sh.getRange(linha, cStatus).setValue(`Atualizado em: ${agora}`);
          sh.getRange(linha, cAgenda).setValue(nomeAgendaDestino);
        } catch (eSet) {
          if (cStatus) {
            const msg = _isPermissionError_(eSet) ? _msgSemPermissao_('atualizar', nomeAgendaDestino) : ('Erro ao atualizar: ' + eSet.message);
            sh.getRange(linha, cStatus).setValue(msg);
          }
        }
        return;
      }
    } catch (e) {
      // getEventById raramente lança; se lançar, tratamos abaixo
    }

    // 2) Não achou no destino → localizar a agenda de ORIGEM
    const origem = localizarEventoEmQualquerAgenda_(idEvento);
    if (!origem) {
      if (cStatus) sh.getRange(linha, cStatus).setValue('Erro: ID não encontrado em nenhuma agenda.');
      return;
    }

    // Se destino == origem, tenta atualizar novamente
    if (origem.id === destino.id) {
      try {
        const ev = origem.cal.getEventById(idEvento) || origem.cal.getEventById(normalizeEventIdForApi_(idEvento));
        if (ev) {
          try {
            ev.setTitle(titulo);
            ev.setLocation(local);
            ev.setDescription(descricao);
            ev.setAllDayDate(dataInicio);
            if (cStatus) sh.getRange(linha, cStatus).setValue(`Atualizado em: ${agora}`);
            sh.getRange(linha, cAgenda).setValue(nomeAgendaDestino);
          } catch (eSet2) {
            if (cStatus) {
              const msg = _isPermissionError_(eSet2) ? _msgSemPermissao_('atualizar', nomeAgendaDestino) : ('Erro ao atualizar: ' + eSet2.message);
              sh.getRange(linha, cStatus).setValue(msg);
            }
          }
          return;
        }
      } catch (_) {}
    }

    // 3) Tenta mover via API avançada; se erro de permissão, cai no fallback
    if (isCalendarApiEnabled_()) {
      const idApi = normalizeEventIdForApi_(idEvento);
      try {
        Calendar.Events.move(origem.id, idApi, destino.id, { sendUpdates: 'all', supportsAttachments: true });
        let eventoDestino = destino.cal.getEventById(idEvento) || destino.cal.getEventById(idApi);
        if (eventoDestino) {
          try {
            eventoDestino.setTitle(titulo);
            eventoDestino.setLocation(local);
            eventoDestino.setDescription(descricao);
            eventoDestino.setAllDayDate(dataInicio);
          } catch (eUpdAfterMove) {
            if (cStatus) {
              const msg = _isPermissionError_(eUpdAfterMove) ? _msgSemPermissao_('atualizar após mover', nomeAgendaDestino)
                                                             : ('Movido, mas erro ao atualizar: ' + eUpdAfterMove.message);
              sh.getRange(linha, cStatus).setValue(msg);
            }
            sh.getRange(linha, cAgenda).setValue(nomeAgendaDestino);
            return;
          }
        }
        if (cStatus) sh.getRange(linha, cStatus).setValue(`Movido (API) e atualizado em: ${agora}`);
        sh.getRange(linha, cAgenda).setValue(nomeAgendaDestino);
        return;
      } catch (eMove) {
        Logger.log('Falha no Calendar.Events.move, aplicando fallback: ' + eMove.message);
      }
    }

    // 4) Fallback: recriar no destino e apagar original
    try {
      const evOrigem = origem.cal.getEventById(idEvento) || origem.cal.getEventById(normalizeEventIdForApi_(idEvento));
      if (!evOrigem) {
        if (cStatus) sh.getRange(linha, cStatus).setValue('Erro: evento sumiu da origem.');
        return;
      }

      let novo;
      try {
        if (evOrigem.isAllDayEvent()) {
          novo = destino.cal.createAllDayEvent(titulo, dataInicio, { location: local, description: descricao });
        } else {
          novo = destino.cal.createEvent(titulo, evOrigem.getStartTime(), evOrigem.getEndTime(), { location: local, description: descricao });
        }
      } catch (eCreateFallback) {
        if (cStatus) {
          const msg = _isPermissionError_(eCreateFallback) ? _msgSemPermissao_('criar (fallback)', nomeAgendaDestino)
                                                           : ('Erro no fallback (criar): ' + eCreateFallback.message);
          sh.getRange(linha, cStatus).setValue(msg);
        }
        return;
      }

      // Se criou no destino, tenta apagar o original
      try {
        evOrigem.deleteEvent();
      } catch (eDelete) {
        if (cStatus) sh.getRange(linha, cStatus).setValue(`Criado no destino, mas falhou ao apagar o original: ${eDelete.message}`);
        sh.getRange(linha, cId).setValue(novo.getId());
        sh.getRange(linha, cAgenda).setValue(nomeAgendaDestino);
        return;
      }

      // Sucesso no fallback completo
      sh.getRange(linha, cId).setValue(novo.getId());
      if (cStatus) sh.getRange(linha, cStatus).setValue(`Movido (fallback) em: ${agora}`);
      sh.getRange(linha, cAgenda).setValue(nomeAgendaDestino);

    } catch (eFb) {
      if (cStatus) sh.getRange(linha, cStatus).setValue('Erro no fallback: ' + eFb.message);
    }

  } catch (err) {
    const cs = _colStatus_(sh);
    if (cs) sh.getRange(linha, cs).setValue('Erro: ' + err.message);
    Logger.log('Erro _sincronizarLinhaEmSheet_: ' + err.message);
  }
}

/* =========================
 * API pública (menus)
 * ========================= */

/** Sincroniza UMA linha na guia "Eventos" (compat). */
function sincronizarEvento(dados, linha) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Eventos');
  if (!sh) {
    SpreadsheetApp.getUi().alert('Aba "Eventos" não encontrada.');
    return;
  }
  _sincronizarLinhaEmSheet_(sh, linha);
}

/** Sincroniza TODAS as linhas da guia "Eventos". */
function sincronizarTodosEventos() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Eventos');
  if (!sh) {
    SpreadsheetApp.getUi().alert('Aba "Eventos" não encontrada.');
    return;
  }
  const last = sh.getLastRow();
  if (last <= 1) {
    SpreadsheetApp.getUi().alert('Nada para sincronizar.');
    return;
  }

  let count = 0;
  for (let r = 2; r <= last; r++) {
    _sincronizarLinhaEmSheet_(sh, r);
    const cTitulo = _colTitulo_(sh) || 3;
    const titulo = (sh.getRange(r, cTitulo).getValue() || '').toString().trim();
    if (titulo) count++;
    Utilities.sleep(30);
  }
  SpreadsheetApp.getUi().alert(`Sincronização concluída! ${count} linha(s) processada(s).`);
}

/** Sincroniza APENAS as linhas selecionadas na guia "Eventos". */
function sincronizarEventosSelecionados() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('Eventos');
  const sel = ss.getActiveRange();

  if (!sh || !sel) {
    SpreadsheetApp.getUi().alert('Selecione as linhas na aba "Eventos".');
    return;
  }

  const r0 = Math.max(sel.getRow(), 2);
  const r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) {
    SpreadsheetApp.getUi().alert('Nada para sincronizar (provavelmente apenas o cabeçalho foi selecionado).');
    return;
  }

  let count = 0;
  for (let r = r0; r <= r1; r++) {
    _sincronizarLinhaEmSheet_(sh, r);
    const cTitulo = _colTitulo_(sh) || 3;
    const titulo = (sh.getRange(r, cTitulo).getValue() || '').toString().trim();
    if (titulo) count++;
    Utilities.sleep(30);
  }
  SpreadsheetApp.getUi().alert(`Sincronização concluída! ${count} linha(s) processada(s).`);
}

/** === Sincroniza TODAS as linhas da GUIA ATUAL (qualquer nome) === */
function sincronizarTodosGuiaAtual() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh) return;
  const last = sh.getLastRow();
  if (last <= 1) {
    SpreadsheetApp.getUi().alert('Nada para sincronizar nesta guia.');
    return;
  }

  // Verifica se a guia tem o conjunto mínimo de colunas
  const cData = _colDataInicio_(sh), cTitulo = _colTitulo_(sh), cId = _colId_(sh), cAgenda = _colAgenda_(sh);
  if (!cData || !cTitulo || !cId || !cAgenda) {
    SpreadsheetApp.getUi().alert('Esta guia não tem o cabeçalho esperado (Data/Título/ID/Agenda).');
    return;
  }

  let count = 0;
  for (let r = 2; r <= last; r++) {
    _sincronizarLinhaEmSheet_(sh, r);
    const titulo = (sh.getRange(r, cTitulo).getValue() || '').toString().trim();
    if (titulo) count++;
    Utilities.sleep(25);
  }
  SpreadsheetApp.getUi().alert(`Sincronização (guia atual) concluída! ${count} linha(s) processada(s).`);
}

/** === Sincroniza APENAS as linhas selecionadas da GUIA ATUAL === */
function sincronizarSelecionadosGuiaAtual() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getActiveSheet();
  const sel = ss.getActiveRange();
  if (!sh || !sel) {
    SpreadsheetApp.getUi().alert('Selecione as linhas na guia atual.');
    return;
  }

  const cData = _colDataInicio_(sh), cTitulo = _colTitulo_(sh), cId = _colId_(sh), cAgenda = _colAgenda_(sh);
  if (!cData || !cTitulo || !cId || !cAgenda) {
    SpreadsheetApp.getUi().alert('Esta guia não tem o cabeçalho esperado (Data/Título/ID/Agenda).');
    return;
  }

  const r0 = Math.max(sel.getRow(), 2);
  const r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) {
    SpreadsheetApp.getUi().alert('Nada para sincronizar (provavelmente selecionou o cabeçalho).');
    return;
  }

  let count = 0;
  for (let r = r0; r <= r1; r++) {
    _sincronizarLinhaEmSheet_(sh, r);
    const titulo = (sh.getRange(r, _colTitulo_(sh)).getValue() || '').toString().trim();
    if (titulo) count++;
    Utilities.sleep(25);
  }
  SpreadsheetApp.getUi().alert(`Sincronização (seleção, guia atual) concluída! ${count} linha(s) processada(s).`);
}

/** Apaga eventos futuros a partir de hoje no calendário do ID_CALENDARIO (compat v1, não usa Agenda). */
function apagarTodosEventos() {
  initCalendario();
  const calendario = CalendarApp.getCalendarById(ID_CALENDARIO);
  const dataInicio = new Date();          // Hoje
  const dataFim = new Date('2030-12-31'); // Futuro arbitrário
  const eventos = calendario.getEvents(dataInicio, dataFim);

  Logger.log(`Encontrados ${eventos.length} eventos a partir de hoje (${dataInicio.toLocaleDateString()}) para apagar no calendário ${ID_CALENDARIO}`);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Eventos');

  eventos.forEach((evento, index) => {
    try {
      const eventoId = evento.getId();
      evento.deleteEvent();
      Logger.log(`Evento ${index + 1}/${eventos.length} apagado - ID: ${eventoId}`);
      const linha = index + 2; // não garante correspondência precisa
      if (sh) sh.getRange(linha, _colId_(sh) || 6).setValue('');
    } catch (erro) {
      Logger.log(`Erro ao apagar evento ${index + 1}: ${erro.message}`);
    }
  });

  SpreadsheetApp.getUi().alert(`Concluído! ${eventos.length} eventos a partir de hoje foram apagados do calendário.`);
}

/** Apaga os eventos das LINHAS selecionadas na guia "Eventos". */
function apagarEventosSelecionados() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName('Eventos');
  const ui  = SpreadsheetApp.getUi();
  if (!sh) return ui.alert('Aba "Eventos" não encontrada.');
  const sel = ss.getActiveRange();
  if (!sel)  return ui.alert('Selecione as LINHAS que deseja apagar.');

  const cId     = _colId_(sh)     || 6;
  const cStatus = _colStatus_(sh) || 7;
  const cAgenda = _colAgenda_(sh) || 8;

  const r0 = Math.max(sel.getRow(), 2);
  const r1 = sel.getRow() + sel.getNumRows() - 1;
  if (r0 > r1) return ui.alert('Nada para processar.');

  let apagados = 0, semId = 0, naoEncontrados = 0, erros = 0;

  for (let r = r0; r <= r1; r++) {
    try {
      const idEventoRaw = (sh.getRange(r, cId).getValue() || '').toString().trim();
      if (!idEventoRaw) {
        semId++;
        if (cStatus) sh.getRange(r, cStatus).setValue('Sem ID para apagar.');
        continue;
      }

      const valorAgenda = (sh.getRange(r, cAgenda).getValue() || '').toString().trim();
      let apagou = false;

      // 1) Tenta apagar na agenda informada
      if (valorAgenda) {
        try {
          const destino = resolverAgenda_(valorAgenda);
          const ev = destino.cal.getEventById(idEventoRaw) ||
                     destino.cal.getEventById(normalizeEventIdForApi_(idEventoRaw));
          if (ev) { ev.deleteEvent(); apagou = true; }
        } catch (_) {}
      }

      // 2) Se não apagou, busca em todas
      if (!apagou) {
        const origem = localizarEventoEmQualquerAgenda_(idEventoRaw);
        if (origem && origem.cal) {
          const ev = origem.cal.getEventById(idEventoRaw) ||
                     origem.cal.getEventById(normalizeEventIdForApi_(idEventoRaw));
          if (ev) { ev.deleteEvent(); apagou = true; }
        }
      }

      if (apagou) {
        apagados++;
        sh.getRange(r, cId).clearContent();
        if (cStatus) sh.getRange(r, cStatus).setValue('Evento apagado do calendário.');
      } else {
        naoEncontrados++;
        if (cStatus) sh.getRange(r, cStatus).setValue('Erro: ID não encontrado.');
      }

      Utilities.sleep(25);
    } catch (e) {
      erros++;
      if (cStatus) sh.getRange(r, cStatus).setValue('Erro ao apagar: ' + e.message);
    }
  }

  ui.alert(`Concluído.
Apagados: ${apagados}
Sem ID: ${semId}
Não encontrados: ${naoEncontrados}
Erros: ${erros}`);
}

/* =========================
 * Busca evento em qualquer agenda
 * ========================= */
function localizarEventoEmQualquerAgenda_(eventId) {
  const cals = CalendarApp.getAllCalendars();
  for (let i = 0; i < cals.length; i++) {
    try {
      const ev = cals[i].getEventById(eventId);
      if (ev) return { cal: cals[i], id: cals[i].getId(), nome: cals[i].getName() };
    } catch(_) {}
  }
  // tenta com ID normalizado (sem @google.com)
  const idApi = normalizeEventIdForApi_(eventId);
  if (idApi !== eventId) {
    for (let i = 0; i < cals.length; i++) {
      try {
        const ev = cals[i].getEventById(idApi);
        if (ev) return { cal: cals[i], id: cals[i].getId(), nome: cals[i].getName() };
      } catch(_) {}
    }
  }
  return null;
}

/* =========================
 * Helpers internos (API)
 * ========================= */
function isCalendarApiEnabled_() {
  try {
    return typeof Calendar !== 'undefined' &&
           Calendar.Events &&
           typeof Calendar.Events.move === 'function';
  } catch (_) {
    return false;
  }
}
function normalizeEventIdForApi_(id) {
  const s = String(id || '').trim();
  return s.endsWith('@google.com') ? s.replace('@google.com', '') : s;
}

/** ======= SHIMS DE COMPATIBILIDADE ======= */
function obterCalendarioPorValorAgenda_(valorAgenda) {
  const res = resolverAgenda_(valorAgenda || '');
  return res && res.cal ? res.cal : CalendarApp.getDefaultCalendar();
}
function obterIdCalendarioPorValorAgenda_(valorAgenda) {
  const res = resolverAgenda_(valorAgenda || '');
  return res && res.id ? res.id : CalendarApp.getDefaultCalendar().getId();
}

/* =========================
 * Testes de permissão (menus)
 * ========================= */

/** Prompt para testar 1 agenda por nome/ID. Mostra alerta com OK / Sem permissão. */
function testarPermissaoAgendaPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Testar permissão em uma agenda', 'Digite o NOME (conforme guia Agendas) ou o ID:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const alvo = (resp.getResponseText() || '').trim();
  if (!alvo) return ui.alert('Nada informado.');

  const res = resolverAgenda_(alvo);
  if (!res || !res.cal) return ui.alert('Agenda não encontrada (verifique nome/ID).');

  const teste = testarPermissaoEscritaCalendario_(res.cal);
  ui.alert(`Agenda: ${res.nome || res.id}\nResultado: ${teste.msg}`);
}

/** Varre a guia "Agendas" (A: Nome, B: ID) e escreve em C: Permissão (OK / Sem permissão / Erro) */
function testarPermissoesGuiaAgendas() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Agendas');
  if (!sh) return SpreadsheetApp.getUi().alert('Guia "Agendas" não encontrada.');

  const last = sh.getLastRow();
  if (last < 2) return SpreadsheetApp.getUi().alert('Guia "Agendas" sem dados.');

  // Cabeçalho da coluna C
  sh.getRange(1, 3).setValue('Permissão');

  const linhas = sh.getRange(2, 1, last - 1, 2).getValues();
  const saida  = [];
  for (let i = 0; i < linhas.length; i++) {
    const nome = (linhas[i][0] || '').toString().trim();
    const id   = (linhas[i][1] || '').toString().trim();
    let msg = '';
    try {
      const alvo = id || nome;
      const res = resolverAgenda_(alvo);
      if (!res || !res.cal) {
        msg = 'Erro: não encontrada';
      } else {
        const t = testarPermissaoEscritaCalendario_(res.cal);
        msg = t.ok ? 'OK' : t.msg;
      }
    } catch (e) {
      msg = 'Erro: ' + e.message;
    }
    saida.push([msg]);
    Utilities.sleep(30);
  }
  sh.getRange(2, 3, saida.length, 1).setValues(saida);
  SpreadsheetApp.getUi().alert('Teste de permissões concluído. Verifique a coluna C da guia "Agendas".');
}

/* =========================
 * Necessita de Utils.gs: formatarDataCompleta(data)
 * ========================= */
// Se não existir no Utils.gs, descomente:
// function formatarDataCompleta(data) {
//   const ano = data.getFullYear();
//   const mes = String(data.getMonth() + 1).padStart(2, '0');
//   const dia = String(data.getDate()).padStart(2, '0');
//   const horas = String(data.getHours()).padStart(2, '0');
//   const minutos = String(data.getMinutes()).padStart(2, '0');
//   const segundos = String(data.getSeconds()).padStart(2, '0');
//   return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
// }
