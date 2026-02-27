// v2 — lê A..H para usar a agenda por linha

function sincronizarEventosSelecionados() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName('Eventos'); // Nome da aba diretamente
  const selecao = planilha.getActiveRange(); // Intervalo selecionado pelo usuário
  const primeiraLinha = selecao.getRow(); // Primeira linha da seleção
  const numLinhas = selecao.getNumRows(); // Número de linhas selecionadas
  let eventosSincronizados = 0;

  // Verifica se a aba selecionada é "Eventos"
  if (aba.getName() !== 'Eventos') {
    SpreadsheetApp.getUi().alert('Erro: Selecione linhas na aba "Eventos".');
    Logger.log(`Aba selecionada (${aba.getName()}) não é Eventos.`);
    return;
  }

  // Verifica se a seleção inclui o cabeçalho (linha 1)
  if (primeiraLinha === 1) {
    SpreadsheetApp.getUi().alert('Erro: Não selecione o cabeçalho (linha 1).');
    Logger.log('Seleção inclui o cabeçalho (linha 1). Saindo.');
    return;
  }

  // Obtém os dados das linhas selecionadas (colunas A a H, 1 a 8)
  const dados = aba.getRange(primeiraLinha, 1, numLinhas, 8).getValues();

  // Processa cada linha selecionada
  dados.forEach((linhaDados, index) => {
    const linha = primeiraLinha + index; // Calcula o número da linha na planilha
    sincronizarEvento(linhaDados, linha);
    if (linhaDados[2]) eventosSincronizados++; // Conta apenas linhas com título (coluna 3)
  });

  SpreadsheetApp.getUi().alert(`Sincronização concluída! ${eventosSincronizados} eventos sincronizados.`);
}
