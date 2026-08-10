// =====================================================
// Apps Script – Rumo Acordo do Orientando
// Conta: faleconosco@institutorumo.com
// Planilha: Acordo do Orientando Rumo
// =====================================================

function doGet(e)  { return processarDados(e); }
function doPost(e) { return processarDados(e); }

function processarDados(e) {
  try {
    var raw = (e.parameter && e.parameter.payload)
           || (e.postData  && e.postData.contents);
    if (!raw) throw new Error('Sem payload');
    var d = JSON.parse(raw);

    // ── Planilha ──────────────────────────────────
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Acordos');

    if (!sheet) {
      sheet = ss.insertSheet('Acordos');

      var cabecalho = [
        'Data e Hora do Aceite',
        'Nome do Orientando',
        'E-mail do Orientando'
      ];

      sheet.appendRow(cabecalho);

      var headerRange = sheet.getRange(1, 1, 1, cabecalho.length);
      headerRange.setBackground('#3B503F');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setFontFamily('Arial');
      headerRange.setFontSize(10);
      sheet.setFrozenRows(1);

      sheet.setColumnWidth(1, 200);
      sheet.setColumnWidth(2, 260);
      sheet.setColumnWidth(3, 260);
    }

    sheet.appendRow([
      d.dataAceite       || '',
      d.nomeOrientando   || '',
      d.emailOrientando  || ''
    ]);

    // ── E-mail de confirmação ─────────────────────
    if (d.emailOrientando) {
      var nome = d.nomeOrientando || 'Orientando(a)';
      var data = d.dataAceite || '';

      var assunto = 'Seu compromisso com o Rumo está registrado 🌱';

      var corpo =
        'Olá, ' + nome + '!\n\n' +
        'Seu aceite ao Acordo do Orientando foi registrado em ' + data + '.\n\n' +
        'Aqui está o que você se comprometeu:\n\n' +
        '✓ Participar do processo com presença, honestidade e abertura\n\n' +
        '✓ Avisar com pelo menos 24 horas de antecedência se precisar cancelar ou remarcar uma sessão — e na sexta-feira anterior, se sua sessão for na segunda-feira\n\n' +
        '✓ Realizar as atividades entre as sessões — pesquisas, leituras, vídeos, podcasts e entrevistas com profissionais — antes do próximo encontro\n\n' +
        'Estamos felizes em ter você nessa jornada. 🌱\n\n' +
        'Instituto Rumo — Orientação Profissional\n' +
        'faleconosco@institutorumo.com';

      MailApp.sendEmail({
        to:      d.emailOrientando,
        subject: assunto,
        body:    corpo
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
