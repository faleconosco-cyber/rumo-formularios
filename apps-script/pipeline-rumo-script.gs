function configurarPipeline() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName("Pipeline");

  // Cabeçalhos
  // Colunas: 1-Data do Contato | 2-Mês de Entrada | 3-Nome Responsável | 4-Nome Adolescente
  //          5-Como Chegou | 6-Telefone | 7-Tipo de Decisão | 8-Triagem Marcada?
  //          9-Data da Triagem | 10-Triagem Realizada? | 11-Data da Proposta | 12-Fechou?
  //          13-Status | 14-Processo | 15-Forma de Pagamento | 16-Valor
  //          17-Próxima Ação | 18-Data da Próxima Ação | 19-Observações
  var headers = [
    "Data do Contato",
    "Mês de Entrada",
    "Nome do Responsável",
    "Nome do Adolescente",
    "Como Chegou",
    "Telefone / WhatsApp",
    "Tipo de Decisão",
    "Triagem Marcada?",
    "Data da Triagem",
    "Triagem Realizada?",
    "Data da Proposta",
    "Fechou?",
    "Status",
    "Processo",
    "Forma de Pagamento",
    "Valor",
    "Próxima Ação",
    "Data da Próxima Ação",
    "Observações"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formatação do cabeçalho
  var cab = sheet.getRange(1, 1, 1, headers.length);
  cab.setBackground("#3E658E");
  cab.setFontColor("#FFFFFF");
  cab.setFontWeight("bold");
  cab.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  // Menus suspensos (linhas 2 a 100)

  // Col 5 — Como Chegou
  sheet.getRange(2, 5, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Indicação", "Google", "Instagram"], true).build());

  // Col 7 — Tipo de Decisão
  sheet.getRange(2, 7, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Individual", "Compartilhada"], true).build());

  // Col 8 — Triagem Marcada?
  sheet.getRange(2, 8, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Sim", "Não"], true).build());

  // Col 10 — Triagem Realizada?
  sheet.getRange(2, 10, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Sim", "Não"], true).build());

  // Col 12 — Fechou?
  sheet.getRange(2, 12, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Sim", "Não"], true).build());

  // Col 13 — Status
  sheet.getRange(2, 13, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Ativo", "Pausado", "Perdido"], true).build());

  // Col 14 — Processo
  sheet.getRange(2, 14, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Cláudia", "Psi do Time"], true).build());

  // Col 15 — Forma de Pagamento
  sheet.getRange(2, 15, 99).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pix à Vista", "Pix Parcelado", "Cartão"], true).build());

  SpreadsheetApp.getUi().alert("Pipeline atualizado com sucesso!");
}
