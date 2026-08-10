// =====================================================
// Apps Script – Rumo Pesquisa de Satisfação
// Conta: faleconosco@institutorumo.com
// Planilha: Pesquisa de Satisfação Rumo
// =====================================================

function doGet(e)  { return processarDados(e); }
function doPost(e) { return processarDados(e); }

function processarDados(e) {
  try {
    var raw = (e.parameter && e.parameter.payload)
           || (e.postData  && e.postData.contents);
    if (!raw) throw new Error('Sem payload');
    var d = JSON.parse(raw);

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Respostas');

    if (!sheet) {
      sheet = ss.insertSheet('Respostas');

      var cabecalho = [
        'Data e Hora',
        'Nome do Responsável',
        'Nome do Filho(a)',
        'Psicóloga',
        'NPS (0–10)',
        'Evolução percebida',
        'Engajamento',
        'Frase do responsável',
        'O que o processo trouxe',
        'Autoriza uso das respostas'
      ];

      sheet.appendRow(cabecalho);

      var headerRange = sheet.getRange(1, 1, 1, cabecalho.length);
      headerRange.setBackground('#3E658E');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setFontFamily('Arial');
      headerRange.setFontSize(10);
      sheet.setFrozenRows(1);

      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 200);
      sheet.setColumnWidth(4, 200);
      sheet.setColumnWidth(5, 80);
      sheet.setColumnWidth(6, 160);
      sheet.setColumnWidth(7, 160);
      sheet.setColumnWidth(8, 320);
      sheet.setColumnWidth(9, 360);
      sheet.setColumnWidth(10, 180);
    }

    var linha = [
      d.dataEnvio        || '',
      d.nomeResponsavel  || '',
      d.nomeFilho        || '',
      d.psicologa        || '',
      d.nps !== undefined ? d.nps : '',
      d.evolucao         || '',
      d.engajamento      || '',
      d.frase            || '',
      d.resultados       || '',
      d.autorizacao      || ''
    ];

    sheet.appendRow(linha);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
