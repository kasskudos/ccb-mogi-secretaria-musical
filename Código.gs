function obterListaMusicalComCookie() {
  var outroEndpointUrl = 'https://musical.congregacao.org.br/grp_musical/listagem';

  // Definir o valor do cookie de sessão
  var cookieParam = 'PHPSESSID=';
  var cookieValue = 'm5qoh833rr0iceb3imptb46far';
  var cookie = cookieParam + cookieValue;

  // Configurar as opções da solicitação para o outro endpoint (também um pedido POST)
  var outroEndpointPayload = {
    // Se houver dados a serem enviados no corpo do POST, adicione-os aqui
    // Exemplo: 'parametro': 'valor'
  };

  var outroEndpointOptions = {
    'method': 'post',
    'headers': {
      'Cookie': cookie
    },
    'payload': outroEndpointPayload
  };

  // Fazer solicitação ao outro endpoint usando o cookie de sessão
  var outroEndpointResponse = UrlFetchApp.fetch(outroEndpointUrl, outroEndpointOptions);

  // Processar a resposta do outro endpoint
  var dadosMusicos = JSON.parse(outroEndpointResponse.getContentText());

  // Chamar a função para adicionar os dados na planilha
  adicionarOuAtualizarDadosNaAba(dadosMusicos);
}

function adicionarOuAtualizarDadosNaAba(dados) {
  // ID da sua planilha
  var spreadsheetId = '1ExhjTg4mEgDspsbTgppD2OYq2juSnumnQJStyF5RWps';

  // Abrir a planilha pelo ID
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // Obter o mês e ano atual
  var hoje = new Date();
  var nomeAba = Utilities.formatDate(hoje, spreadsheet.getSpreadsheetTimeZone(), 'MMM-yyyy'); // Formato JAN-2024

  // Verificar se a aba já existe
  var sheet = spreadsheet.getSheetByName(nomeAba);

  if (sheet) {
    // Limpar o conteúdo a partir da segunda linha
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  } else {
    // Criar uma nova aba com o nome do mês e ano atual
    sheet = spreadsheet.insertSheet(nomeAba);
  }

  // Adicionar os dados à planilha a partir da segunda linha
  sheet.getRange(2, 1, dados.data.length, dados.data[0].length).setValues(dados.data);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({column: 1, ascending: true});
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();
}
