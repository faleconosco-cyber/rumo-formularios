// =====================================================
// Apps Script – Banco de Profissionais Rumo
// Conta: faleconosco@institutorumo.com
// Planilha: Banco de Profissionais Rumo
//
// Duas abas. "Profissionais" recebe uma linha por voluntário, na ordem de
// chegada. "Por área" se organiza sozinha por fórmula, agrupada por área e
// nome, e é a aba de consulta: é dela que sai o contato para passar ao
// adolescente depois que o voluntário topar a conversa.
// =====================================================

function doGet(e)  { return processarDados(e); }
function doPost(e) { return processarDados(e); }

var ABA_BASE  = 'Profissionais';
var ABA_AREA  = 'Por área';

var CABECALHO = [
  'Data e Hora',
  'Nome',
  'Área',
  'Profissão ou cargo',
  'Formação',
  'Tempo na área',
  'O que faz no dia a dia',
  'E-mail',
  'WhatsApp',
  'Responsável por',
  'Consentimento'
];

// A aba de consulta, para achar quem chamar e já ter o contato à mão
var CABECALHO_AREA = ['Área', 'Nome', 'Profissão ou cargo', 'Formação', 'Tempo na área', 'O que faz no dia a dia', 'E-mail', 'WhatsApp'];

function processarDados(e) {
  try {
    var raw = (e.parameter && e.parameter.payload)
           || (e.postData  && e.postData.contents);
    if (!raw) throw new Error('Sem payload');
    var d = JSON.parse(raw);

    var sheet = pegarBase();
    sheet.appendRow([
      d.dataEnvio      || '',
      d.nome           || '',
      d.area           || '',
      d.cargo          || '',
      d.formacao       || '',
      d.tempo          || '',
      d.diaadia        || '',
      d.email          || '',
      d.whatsapp       || '',
      d.orientando     || '',
      d.consentimento  || ''
    ]);

    garantirAbaPorArea();
    confirmarPorEmail(d);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function pegarBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ABA_BASE);
  if (sheet) return sheet;

  sheet = ss.insertSheet(ABA_BASE, 0);
  sheet.appendRow(CABECALHO);
  estilizarCabecalho(sheet, CABECALHO.length);

  sheet.setColumnWidth(1, 150);  // data
  sheet.setColumnWidth(2, 210);  // nome
  sheet.setColumnWidth(3, 220);  // área
  sheet.setColumnWidth(4, 210);  // cargo
  sheet.setColumnWidth(5, 210);  // formação
  sheet.setColumnWidth(6, 140);  // tempo
  sheet.setColumnWidth(7, 420);  // dia a dia
  sheet.setColumnWidth(8, 230);  // e-mail
  sheet.setColumnWidth(9, 160);  // whatsapp
  sheet.setColumnWidth(10, 200); // responsável por
  sheet.setColumnWidth(11, 260); // consentimento
  return sheet;
}

// Aba de consulta, agrupada por área. Só precisa ser montada uma vez;
// depois a fórmula se atualiza sozinha a cada novo voluntário.
// Use esta aba para achar quem chamar. O contato só sai daqui depois que o
// voluntário disser que pode participar daquela conversa.
function garantirAbaPorArea() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(ABA_AREA)) return;

  var sheet = ss.insertSheet(ABA_AREA);
  sheet.appendRow(CABECALHO_AREA);
  estilizarCabecalho(sheet, CABECALHO_AREA.length);

  sheet.getRange('A2').setFormula(
    '=IFERROR(QUERY(' + ABA_BASE + '!A2:K, ' +
    '"select C, B, D, E, F, G, H, I where C is not null order by C, B", 0), )'
  );

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 210);
  sheet.setColumnWidth(3, 210);
  sheet.setColumnWidth(4, 210);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 420);
  sheet.setColumnWidth(7, 230);
  sheet.setColumnWidth(8, 160);
  sheet.getRange('F:F').setWrap(true);
}

function estilizarCabecalho(sheet, colunas) {
  var faixa = sheet.getRange(1, 1, 1, colunas);
  faixa.setBackground('#3B503F');
  faixa.setFontColor('#FFFFFF');
  faixa.setFontWeight('bold');
  faixa.setFontFamily('Arial');
  faixa.setFontSize(10);
  sheet.setFrozenRows(1);
}

// Confirmação para o voluntário, que também serve de registro do combinado
function confirmarPorEmail(d) {
  if (!d.email) return;
  var primeiro = (d.nome || '').split(' ')[0];

  MailApp.sendEmail({
    to: d.email,
    subject: 'Você está no Banco de Profissionais Rumo 🌱',
    body:
      'Oi, ' + primeiro + '!\n\n' +
      'Obrigada por entrar no Banco de Profissionais do Instituto Rumo. ' +
      'Adolescente que conversa com gente de verdade escolhe com muito menos medo, e agora você faz parte disso.\n\n' +
      'Foi assim que ficou registrado:\n\n' +
      '• Área: ' + (d.area || '') + '\n' +
      '• Profissão: ' + (d.cargo || '') + '\n' +
      '• Formação: ' + (d.formacao || '') + '\n' +
      '• Tempo na área: ' + (d.tempo || '') + '\n\n' +
      'O combinado, para você guardar:\n\n' +
      '• Eu entro em contato antes de cada conversa, e você decide se pode ou não participar.\n' +
      '• Se você topar, eu passo o seu contato para o adolescente e vocês combinam entre vocês o que for melhor para os dois.\n' +
      '• A conversa é online ou por chamada de voz, pelo tempo que você tiver disponível.\n' +
      '• Para sair do banco, responda este e-mail. Saída imediata, sem justificar nada.\n\n' +
      'Qualquer coisa é só chamar.\n\n' +
      'Cláudia Botelho\n' +
      'Psicóloga CRP 05/73536\n' +
      'Instituto Rumo\n' +
      'faleconosco@institutorumo.com'
  });
}
