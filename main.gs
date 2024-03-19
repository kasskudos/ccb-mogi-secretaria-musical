var COOKIE = '2gqfao56dlu67qcltbdruk6g3n';
var SHEET_ID = '1ExhjTg4mEgDspsbTgppD2OYq2juSnumnQJStyF5RWps';

// Open the spreadsheet by ID
var spreadsheet = SpreadsheetApp.openById(SHEET_ID);

function fetchAndProcessMusicalList() {
  var endpointUrl = 'https://musical.congregacao.org.br/grp_musical/listagem';

  var endpointOptions = {
    'method': 'post',
    'headers': {
      'Cookie': `PHPSESSID=${COOKIE}` // Set the session cookie value
    }
  };

  // Make a request to the musical endpoint using the session cookie
  var endpointResponse = UrlFetchApp.fetch(endpointUrl, endpointOptions);

  // Process the response from the musical endpoint
  var musiciansData = JSON.parse(endpointResponse.getContentText());



  // Get the current month and year
  var today = new Date();
  var formattedSheetName = Utilities.formatDate(today, spreadsheet.getSpreadsheetTimeZone(), 'MMM-yyyy');

  // Call the function to add the data to the spreadsheet
  addOrUpdateDataToSheet(musiciansData,formattedSheetName);
}

function fetchAndProcessChurchList() {
  var churchEndpointUrl = 'https://musical.congregacao.org.br/igrejas/listagem';

  var churchEndpointOptions = {
    'method': 'post',
    'headers': {
      'Cookie': `PHPSESSID=${COOKIE}` // Set the session cookie value
    }
  };

  // Make a request to the church endpoint using the session cookie
  var churchEndpointResponse = UrlFetchApp.fetch(churchEndpointUrl, churchEndpointOptions);

  // Process the response from the church endpoint
  var churchData = JSON.parse(churchEndpointResponse.getContentText());

  // Call the function to add the church data to a new sheet
  addOrUpdateDataToSheet(churchData, 'Lista de Igrejas');
}

function fetchAndUpdateTotalMusicians() {
  // Obter a aba "Lista de Igrejas"
  var sheet = spreadsheet.getSheetByName('Lista de Igrejas');

  // Verificar se a aba existe
  if (!sheet) {
    Logger.log("A aba 'Lista de Igrejas' não foi encontrada na planilha.");
    return;
  }

  // Obter os dados da coluna A (ID da igreja), ignorando a primeira linha (cabeçalho)
  var churchIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();

  // Fazer uma solicitação para cada ID de igreja
  for (var i = 0; i < churchIds.length; i++) {
    var churchId = churchIds[i];

    if (churchId) {
      var totalMusicians = getTotalMusiciansForChurch(churchId);
      var totalAlunos = getNumberOfCandidatesForChurch(churchId);

      // Atualizar a planilha com o total de músicos na última coluna (coluna K)      
      sheet.getRange(i + 2, 10).setValue(totalAlunos); // 10 representa a coluna J
      sheet.getRange(i + 2, 11).setValue(totalMusicians); // 11 representa a coluna K

    }
  }
}

function getTotalMusiciansForChurch(churchId) {
  var endpointUrl = 'https://musical.congregacao.org.br/equilibrio_orquestral/total_musicos';

  var endpointPayload = {
    'id_igreja': churchId
  };

  var endpointOptions = {
    'method': 'post',
    'headers': {
      'Cookie': `PHPSESSID=${COOKIE}`
    },
    'payload': endpointPayload
  };

  // Fazer a solicitação para obter o total de músicos
  var endpointResponse = UrlFetchApp.fetch(endpointUrl, endpointOptions);

  // Processar a resposta e retornar o total de músicos
  var total_musicos = JSON.parse(endpointResponse.getContentText());
  return total_musicos;
}

function getTotalCandidates(htmlContent) {
  // Utilizando RegExp para encontrar a parte do HTML que contém o número de candidatos
  var regex = /CANDIDATO\(A\)/gi;
  var matches = htmlContent.match(regex);

  // Retornar o número total de candidatos
  return matches ? matches.length : 0;
}

function getNumberOfCandidatesForChurch(churchId) {
  var endpointUrl = 'https://musical.congregacao.org.br/igrejas/editar/' + churchId;

  var endpointOptions = {
    'method': 'get',
    'headers': {
      'Cookie': `PHPSESSID=${COOKIE}`
    }
  };

  // Fazer uma solicitação GET para obter o HTML da página
  var endpointResponse = UrlFetchApp.fetch(endpointUrl, endpointOptions);

  // Extrair o número de alunos do conteúdo HTML
  var totalAlunos = getTotalCandidates(endpointResponse.getContentText());

  // Retornar o número total de alunos
  return totalAlunos;
}

function addOrUpdateDataToSheet(data, sheetName) {
  // Check if the sheet already exists
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    // Clear content from the second row onward
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  } else {
    // Create a new sheet with the provided name
    sheet = spreadsheet.insertSheet(sheetName);
  }

  // Add data to the sheet starting from the second row
  sheet.getRange(2, 1, data.data.length, data.data[0].length).setValues(data.data);

  // Sort the data by the ID column (column 1), starting from the second row
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({column: 1, ascending: true});

  // Check if a filter already exists before creating a new one
  var existingFilter = sheet.getFilter();
  if (!existingFilter) {
    // Add a filter to the sheet
    sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();
  }
}