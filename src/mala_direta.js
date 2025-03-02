function malaDireta() {
  try {
    Logger.log("Iniciando função malaDireta...");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var planilha = ss.getActiveSheet();
    var dados = planilha.getDataRange().getValues();
    Logger.log("Dados carregados da planilha.");

  // Colunas
  var COLUNA_STATUS = 12;  // M
  var COLUNA_NOME = 15;    // P
  var COLUNA_ENDERECO = 24;// Y
  var COLUNA_CIDADE = 27;  // AB
  var COLUNA_CEP = 25;     // Z
  var COLUNA_INSTITUICAO = 5; // F
  var COLUNA_PRONOME_ABR = 17; // R
  var COLUNA_SEXO = 14; // O
  var COLUNA_STATUS_ATUALIZATION = 13;

  var registrosNaoImpressos = [];

  var modeloId = "1YIPwCWjSDsRF7BJ6vqfwjDs1FwzVZ-_iddy2Jkfe43w";
  var timestamp = Utilities.formatDate(new Date(), "GMT", "yyyyMMddHHmmss");
  var docCopy = DriveApp.getFileById(modeloId).makeCopy("Mala Direta ENVELOPE DL " + timestamp);
  var corpoDoc = DocumentApp.openById(docCopy.getId()).getBody();

    for (var i = 2; i < dados.length; i++) {
    if (dados[i][COLUNA_STATUS] == "PARA IMPRESSÃO") {

        corpoDoc.appendParagraph("");

        var tratamento = (dados[i][COLUNA_SEXO] == "Masculino") ? "Ao" : "À";

        // Adicionando pronome, nome e dando destaque
        corpoDoc.appendParagraph(tratamento + " " + (dados[i][COLUNA_PRONOME_ABR] || "") + " " + (dados[i][COLUNA_NOME] || "")).setBold(true);

        // Adicionando a instituição, se disponível
        if (dados[i][COLUNA_INSTITUICAO]) {
            corpoDoc.appendParagraph(dados[i][COLUNA_INSTITUICAO]);
        } else {
            corpoDoc.appendParagraph("-");
        }

        // Adicionando endereço
        corpoDoc.appendParagraph(dados[i][COLUNA_ENDERECO] || "-");

        // Adicionando CEP e cidade
        var cepCidade = (dados[i][COLUNA_CEP] || "") + " - " + (dados[i][COLUNA_CIDADE] || "");
        corpoDoc.appendParagraph(cepCidade);

        corpoDoc.appendPageBreak();

        // Atualiza o status na planilha
        planilha.getRange(i + 1, COLUNA_STATUS_ATUALIZATION).setValue("MALA DIRETA PRONTA");
        Logger.log("Atualizado o status da linha " + (i + 1) + " para 'MALA DIRETA PRONTA'.");
    }
}


// Fecha o documento para salvar as alterações
var doc = DocumentApp.openById(docCopy.getId());
doc.saveAndClose();
Logger.log("Documento salvo e fechado.");





  var fileUrl = docCopy.getUrl();

    if (registrosNaoImpressos.length > 0) {
      Logger.log("Registros não impressos devido à falta de dados: " + registrosNaoImpressos.join(", "));
      SpreadsheetApp.getUi().alert("Atenção! Os seguintes registros não foram impressos devido à falta de dados: \n\n" + registrosNaoImpressos.join(", "));
    } else {
      Logger.log("Operação concluída com sucesso! Todos os registros foram processados.");
      SpreadsheetApp.getUi().alert("Operação concluída com sucesso! Todos os registros foram processados.");
    }
    Logger.log("Mala Direta concluída com sucesso.");
    
    var htmlOutput = HtmlService
      .createHtmlOutput('Documento criado com sucesso!\n<a href="' + fileUrl + '" target="_blank">' + fileUrl + '</a>')
      .setWidth(350)
      .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Script concluído');

  } catch (error) {
    Logger.log("Erro na função malaDireta: " + error.toString());
    // Captura qualquer erro ocorrido durante a execução e apresenta ao usuário
    SpreadsheetApp.getUi().alert("Ocorreu um erro durante a execução: " + error.toString());
  }
}

function addMenu() {
  try {
    Logger.log("Iniciando função addMenu...");
    var menu = SpreadsheetApp.getUi().createMenu('Exportar');
    menu.addItem('IMPRESSÃO ENVELOPE DL', 'malaDireta');
    menu.addToUi();
    Logger.log("Menu adicionado com sucesso!");
  } catch (error) {
    SpreadsheetApp.getUi().alert("Ocorreu um erro ao adicionar o menu: " + error.toString());
  }
}

function onOpen(e) {
  try {
    Logger.log("Iniciando função onOpen...");
    addMenu();
  Logger.log("Script iniciado com sucesso.");
  } catch (error) {
    Logger.log("Erro ao iniciar o script: " + error.toString());
    SpreadsheetApp.getUi().alert("Ocorreu um erro ao iniciar o script: " + error.toString());
  }
}
