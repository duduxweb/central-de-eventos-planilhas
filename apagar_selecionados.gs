// v2 — Apaga do Google Calendar por linha, respeitando a agenda informada na coluna H

/**
 * Apaga do Google Calendar os eventos correspondentes às linhas selecionadas na planilha.
 * A função verifica a coluna F ("ID do Evento") e a coluna H ("Agenda") de cada linha selecionada.
 * Se H for um NOME, mapeia em "Agendas" (A:Nome, B:ID). Se H for um ID, usa direto.
 * Fallbacks: Configurações!B1 → calendário padrão.
 */
function apagarEventosSelecionados() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getActiveSheet();

  // 1. Valida se a aba ativa é a correta
  if (aba.getName() !== 'Eventos') {
    SpreadsheetApp.getUi().alert('Erro: Por favor, execute esta função com as linhas selecionadas na aba "Eventos".');
    return;
  }

  const selecao = planilha.getActiveRange();
  const primeiraLinha = selecao.getRow();

  // 2. Valida se o cabeçalho foi selecionado
  if (primeiraLinha === 1) {
    SpreadsheetApp.getUi().alert('Erro: Não selecione a linha do cabeçalho (linha 1). A operação foi cancelada.');
    return;
  }

  // 3. Processa cada linha selecionada para apagar o evento
  const numLinhas = selecao.getNumRows();
  const rangeLinhas = aba.getRange(primeiraLinha, 1, numLinhas, 8); // Colunas A até H
  const dados = rangeLinhas.getValues();
  let eventosApagados = 0;

  dados.forEach((linhaDados, index) => {
    const idEvento = (linhaDados[5] || '').toString().trim(); // Coluna F (índice 5)
    const agendaValor = (linhaDados[7] || '').toString().trim(); // Coluna H (índice 7)
    const linhaAtual = primeiraLinha + index;

    if (!idEvento) return;

    try {
      const calendario = obterCalendarioPorValorAgenda_(agendaValor);
      const evento = calendario.getEventById(idEvento);
      if (evento) {
        evento.deleteEvent();
        Logger.log(`Evento com ID ${idEvento} apagado do calendário (linha ${linhaAtual}).`);

        // Limpa o ID e atualiza o status na planilha para dar feedback
        aba.getRange(linhaAtual, 6).clearContent(); // Limpa coluna F
        aba.getRange(linhaAtual, 7).setValue('Evento apagado do calendário.'); // Atualiza coluna G
        eventosApagados++;
      } else {
        aba.getRange(linhaAtual, 7).setValue('Erro: ID não encontrado no calendário.');
        Logger.log(`ID de evento ${idEvento} (linha ${linhaAtual}) não foi encontrado no calendário.`);
      }
    } catch (e) {
      Logger.log(`Erro ao tentar apagar o evento da linha ${linhaAtual} (ID: ${idEvento}): ${e.message}`);
      aba.getRange(linhaAtual, 7).setValue(`Erro ao apagar: ${e.message}`);
    }
  });

  // 4. Exibe um resumo da operação para o usuário
  if (eventosApagados > 0) {
    SpreadsheetApp.getUi().alert(`${eventosApagados} evento(s) selecionado(s) foram apagados do calendário com sucesso!`);
  } else {
    SpreadsheetApp.getUi().alert('Nenhum evento foi apagado. Verifique se as linhas selecionadas possuem um "ID do Evento" válido na coluna F.');
  }
}
